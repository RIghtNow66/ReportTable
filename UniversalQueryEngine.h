#pragma once
#pragma once
#ifndef UNIVERSALQUERYENGINE_H
#define UNIVERSALQUERYENGINE_H

#include <map>
#include <vector>
#include <string>
#include <QVariant>
#include <QHash>

// ͨ�ò�ѯ����
class  UniversalQueryEngine {
public:
    static UniversalQueryEngine& instance();

    // ִ��RTDB��ѯ
    QHash<QString, QVariant> queryValuesForBindingKeys(
        QList<QString> bindingKeys
    );

private:
    UniversalQueryEngine() = default;
    ~UniversalQueryEngine() = default;
    UniversalQueryEngine(const UniversalQueryEngine&) = delete;
    UniversalQueryEngine& operator=(const UniversalQueryEngine&) = delete;

    struct RtdbRequest
    {
        QString originalKey;
        std::string tableName;
        std::string objName;
        std::string fieldName;
    };


    // ���������ַ���
    void executeRtdQuery(
        const std::vector<RtdbRequest>& requests,
        QHash<QString, QVariant>& results
    );

};

#endif // UNIVERSALQUERYENGINE_H