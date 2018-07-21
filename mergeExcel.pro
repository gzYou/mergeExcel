#-------------------------------------------------
#
# Project created by QtCreator 2018-06-26T14:31:58
#
#-------------------------------------------------

QT       += core gui
CONFIG   += qaxcontainer

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

TARGET = mergeExcel
TEMPLATE = app


SOURCES += main.cpp\
        mainwindow.cpp \
    excelengine.cpp

HEADERS  += mainwindow.h \
    excelengine.h

FORMS    += mainwindow.ui
