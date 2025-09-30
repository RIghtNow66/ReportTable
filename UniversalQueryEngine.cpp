#include "UniversalQueryEngine.h"
#include "platform/rdbapi.h"  // RTDB���ͷ�ļ�
#include <QRegularExpression>
#include <QDebug>
#include <sstream>

UniversalQueryEngine& UniversalQueryEngine::instance() {
    static UniversalQueryEngine instance;
    return instance;
}

// --- �µĹ����ӿ�ʵ�� ---
QHash<QString, QVariant> UniversalQueryEngine::queryValuesForBindingKeys(QList<QString> bindingKeys)
{
    QHash<QString, QVariant> finalResults;
    if (bindingKeys.isEmpty()) {
        return finalResults;
    }

    // 1. �������룬���������б�
    std::vector<RtdbRequest> requests;
    for (const QString& key : bindingKeys) {
        if (!key.startsWith("##")) continue;

        QString cleanKey = key.mid(2); // ȥ�� "##"
        QStringList parts = cleanKey.split('/');

        if (parts.size() >= 3) {
            RtdbRequest req;
            req.originalKey = key;
            req.tableName = parts[0].toStdString();
            req.objName = parts[1].toStdString();
            req.fieldName = parts[2].toStdString();
            requests.push_back(req);
        }
        else {
            // ���ڸ�ʽ�����Key��ֱ���ڽ���б��
            finalResults[key] = "Invalid Format";
        }
    }

    // 2. ִ�в�ѯ
    if (!requests.empty()) {
        executeRtdQuery(requests, finalResults);
    }

    // 3. �������õĽ��
    return finalResults;
}


void UniversalQueryEngine::executeRtdQuery(
    const std::vector<RtdbRequest>& requests,
    QHash<QString, QVariant>& results // ֱ�������
)
{
    int N = requests.size();
    if (N == 0) return;

    // 1. ׼��RTDB��ѯ������ṹ
    std::vector<RDB_FIELD_STRU> getfinfo(N);
    for (int i = 0; i < N; ++i) {
        // ʹ��C++11�İ�ȫ���� strcpy_s������ȷ���������㹻��������ʹ�� strcpy
        strncpy(getfinfo[i].tabname, requests[i].tableName.c_str(), sizeof(getfinfo[i].tabname) - 1);
        strncpy(getfinfo[i].objname, requests[i].objName.c_str(), sizeof(getfinfo[i].objname) - 1);
        strncpy(getfinfo[i].fldname, requests[i].fieldName.c_str(), sizeof(getfinfo[i].fldname) - 1);
        // ȷ���ַ����� null ��β
        getfinfo[i].tabname[sizeof(getfinfo[i].tabname) - 1] = '\0';
        getfinfo[i].objname[sizeof(getfinfo[i].objname) - 1] = '\0';
        getfinfo[i].fldname[sizeof(getfinfo[i].fldname) - 1] = '\0';
    }

    // 2. ִ��RTDB��ѯ
    Rdb_QuickPolling rsp;
    Rdb_MultiTypeValue rmtv;
    int ret = rsp.RdbGetFieldValue(SYS_USER, "", N, getfinfo.data(), &rmtv);

    // 3. �����ѯ���
    if (ret <= 0) {
        qDebug() << "RdbGetFieldValue API failed, ret =" << ret;
        // ���API����ʧ�ܣ�������������Ϊ����
        for (const auto& req : requests) {
            results[req.originalKey] = "Query Error";
        }
        return;
    }

    // ����һ�����������������Щ�������յ���Ӧ
    std::vector<bool> responded(N, false);

    for (int i = 0; i < ret; i++) {
        int parano = rmtv.RdbGetValOrderno(i); // ��ȡ��Ӧ��Ӧ��ԭʼ��������

        if (parano >= 0 && parano < N) {
            // ��ȡ��Ӧ��ԭʼ����
            const RtdbRequest& originalRequest = requests[parano];

            // ��RTDB����л�ȡֵ (���������int�������Ը�����Ҫʹ�� RdbGetVal_double, RdbGetVal_string ��)
            int resultValue = rmtv.RdbGetVal_int(i);

            // ʹ��ԭʼKey������������ϣ��
            results[originalRequest.originalKey] = QVariant(resultValue);

            // ��Ǵ���������Ӧ
            responded[parano] = true;
        }
    }

    // 4. ����δ�յ���Ӧ������
    for (int i = 0; i < N; ++i) {
        if (!responded[i]) {
            results[requests[i].originalKey] = "No Response";
        }
    }
}