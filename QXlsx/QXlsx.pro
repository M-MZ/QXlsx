QT -= gui

CONFIG += c++11 console
CONFIG -= app_bundle

# The following define makes your compiler emit warnings if you use
# any Qt feature that has been marked deprecated (the exact warnings
# depend on your compiler). Please consult the documentation of the
# deprecated API in order to know how to port your code away from it.
DEFINES += QT_DEPRECATED_WARNINGS

# You can also make your code fail to compile if it uses deprecated APIs.
# In order to do so, uncomment the following line.
# You can also select to disable deprecated APIs only up to a certain version of Qt.
#DEFINES += QT_DISABLE_DEPRECATED_BEFORE=0x060000    # disables all the APIs deprecated before Qt 6.0.0

SOURCES += \
        main.cpp \
    source/xlsxabstractooxmlfile.cpp \
    source/xlsxabstractsheet.cpp \
    source/xlsxcell.cpp \
    source/xlsxcellformula.cpp \
    source/xlsxcelllocation.cpp \
    source/xlsxcellrange.cpp \
    source/xlsxcellreference.cpp \
    source/xlsxchart.cpp \
    source/xlsxchartsheet.cpp \
    source/xlsxcolor.cpp \
    source/xlsxconditionalformatting.cpp \
    source/xlsxcontenttypes.cpp \
    source/xlsxdatavalidation.cpp \
    source/xlsxdocpropsapp.cpp \
    source/xlsxdocpropscore.cpp \
    source/xlsxdocument.cpp \
    source/xlsxdrawing.cpp \
    source/xlsxdrawinganchor.cpp \
    source/xlsxformat.cpp \
    source/xlsxmediafile.cpp \
    source/xlsxnumformatparser.cpp \
    source/xlsxrelationships.cpp \
    source/xlsxrichstring.cpp \
    source/xlsxsharedstrings.cpp \
    source/xlsxsimpleooxmlfile.cpp \
    source/xlsxstyles.cpp \
    source/xlsxtheme.cpp \
    source/xlsxutility.cpp \
    source/xlsxworkbook.cpp \
    source/xlsxworksheet.cpp \
    source/xlsxzipreader.cpp \
    source/xlsxzipwriter.cpp \
    main.cpp

# Default rules for deployment.
qnx: target.path = /tmp/$${TARGET}/bin
else: unix:!android: target.path = /opt/$${TARGET}/bin
!isEmpty(target.path): INSTALLS += target

# QXlsx code for Application Qt project
QXLSX_PARENTPATH=./         # current QXlsx path is . (. means curret directory)
QXLSX_HEADERPATH=./header/  # current QXlsx header path is ./header/
QXLSX_SOURCEPATH=./source/  # current QXlsx source path is ./source/
include(./QXlsx.pri)

HEADERS += \
    header/xlsxabstractooxmlfile.h \
    header/xlsxabstractooxmlfile_p.h \
    header/xlsxabstractsheet.h \
    header/xlsxabstractsheet_p.h \
    header/xlsxcell.h \
    header/xlsxcell_p.h \
    header/xlsxcellformula.h \
    header/xlsxcellformula_p.h \
    header/xlsxcelllocation.h \
    header/xlsxcellrange.h \
    header/xlsxcellreference.h \
    header/xlsxchart.h \
    header/xlsxchart_p.h \
    header/xlsxchartsheet.h \
    header/xlsxchartsheet_p.h \
    header/xlsxcolor_p.h \
    header/xlsxconditionalformatting.h \
    header/xlsxconditionalformatting_p.h \
    header/xlsxcontenttypes_p.h \
    header/xlsxdatavalidation.h \
    header/xlsxdatavalidation_p.h \
    header/xlsxdocpropsapp_p.h \
    header/xlsxdocpropscore_p.h \
    header/xlsxdocument.h \
    header/xlsxdocument_p.h \
    header/xlsxdrawing_p.h \
    header/xlsxdrawinganchor_p.h \
    header/xlsxformat.h \
    header/xlsxformat_p.h \
    header/xlsxglobal.h \
    header/xlsxmediafile_p.h \
    header/xlsxnumformatparser_p.h \
    header/xlsxrelationships_p.h \
    header/xlsxrichstring.h \
    header/xlsxrichstring_p.h \
    header/xlsxsharedstrings_p.h \
    header/xlsxsimpleooxmlfile_p.h \
    header/xlsxstyles_p.h \
    header/xlsxtheme_p.h \
    header/xlsxutility_p.h \
    header/xlsxworkbook.h \
    header/xlsxworkbook_p.h \
    header/xlsxworksheet.h \
    header/xlsxworksheet_p.h \
    header/xlsxzipreader_p.h \
    header/xlsxzipwriter_p.h
