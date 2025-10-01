#pragma once
// TimeSettingsDialog.h
#ifndef TIMESETTINGSDIALOG_H
#define TIMESETTINGSDIALOG_H

#include <QDialog>
#include <QDateTime>
#include <QDateTimeEdit>
#include <QSpinBox>
#include <QComboBox>
#include <QRadioButton>
#include <QGroupBox>
#include <QPushButton>
#include <QLabel>
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QButtonGroup>

class TimeSettingsDialog : public QDialog
{
    Q_OBJECT

public:
    enum ReportType {
        Daily = 0,   // 日报
        Weekly = 1,  // 周报
        Monthly = 2, // 月报
        Custom = 3   // 自定义 - 新增
    };

    explicit TimeSettingsDialog(QWidget* parent = nullptr);

    // 获取用户设置
    QDateTime getStartTime() const;
    QDateTime getEndTime() const;
    int getIntervalSeconds() const;
    ReportType getReportType() const;

    // 设置初始值（用于记忆上次选择）
    void setStartTime(const QDateTime& time);
    void setReportType(ReportType type);

private slots:
    void onReportTypeChanged(int id);
    void onStartTimeChanged(const QDateTime& dateTime);
    void onEndTimeChanged(const QDateTime& dateTime);
    void onIntervalValueChanged(int value);
    void onIntervalUnitChanged(int index);
    void onQuickIntervalClicked();

private:
    void setupUI();
    void calculateEndTime();
    void updateIntervalDisplay();
    QDateTime limitToCurrentTime(const QDateTime& time);

    void updateEndTimeEditability();  // 新增：更新终止时间是否可编辑
	void adjustIntervalForReportType(); // 新增：根据报表类型调整间隔的合理范围和默认值

private:
    // UI组件
    QButtonGroup* m_reportTypeGroup;
    QRadioButton* m_dailyRadio;
    QRadioButton* m_weeklyRadio;
    QRadioButton* m_monthlyRadio;
    QRadioButton* m_customRadio;  // 新增

    QDateTimeEdit* m_startTimeEdit;
    QDateTimeEdit* m_endTimeEdit;

    QSpinBox* m_intervalSpinBox;
    QComboBox* m_intervalUnitCombo;
    QLabel* m_intervalDisplayLabel;

    QPushButton* m_quick5SecBtn;
    QPushButton* m_quick1MinBtn;
    QPushButton* m_quick5MinBtn;
    QPushButton* m_quick1HourBtn;

    QPushButton* m_okBtn;
    QPushButton* m_cancelBtn;

    // 数据
    ReportType m_currentType;
    QDateTime m_startTime;
    QDateTime m_endTime;
    int m_intervalValue;
    int m_intervalUnit; // 0=秒, 1=分钟, 2=小时
};

#endif // TIMESETTINGSDIALOG_H