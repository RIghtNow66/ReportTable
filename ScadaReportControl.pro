TEMPLATE = app
LANGUAGE = C++
TARGET = ScadaReportControl
QT += core widgets

include(QXlsx/QXlsx/QXlsx.pri)

SOURCES += \
    main.cpp\
    mainwindow.cpp\
    reportdatamodel.cpp\
    formulaengine.cpp\
    excelhandler.cpp\
	ReportTable.cpp

HEADERS +=\
    mainwindow.h\
    reportdatamodel.h\
    formulaengine.h\
    excelhandler.h\
    Cell.h\
	ReportTable.h

RESOURCES += ReportTable.qrc
