// main.cpp

#include <QCoreApplication>
#include <QtCore>

#include "xlsxdocument.h"
#include "xlsxchartsheet.h"
#include "xlsxcellrange.h"
#include "xlsxchart.h"
#include "xlsxrichstring.h"
#include "xlsxworkbook.h"
using namespace QXlsx;

enum Type{Scatter, Line, Bar, Pie};
Chart* ProduceChart (Chart *, Type);
Document* AddData(Document*);

int main(int argc, char *argv[])
{
    QCoreApplication a(argc, argv);
    QXlsx::Document xlsx;

    xlsx.addSheet("Testing");
    xlsx.write(1, 1, "Data A");
    xlsx.write(1, 2, "UCL");
    xlsx.write(1, 3, "LCL");
    xlsx.write(1, 4, "Center Line");
    xlsx.write(1, 5, "USL");
    xlsx.write(1, 6, "LSL");

    for (int i=2; i<46; ++i) {
    xlsx.write(i, 1, rand() % 34 + 90); //A1:A9
    xlsx.write(i, 2, 177); //B1:B9
    xlsx.write(i, 3, 47); //C1:C9
    xlsx.write(i, 4, 105); //B1:B9
    xlsx.write(i, 5, 150); //C1:C9
    xlsx.write(i, 6, 50); //C1:C9
    }
    Chart *scatterChart = xlsx.insertChart(3, 7, QSize(800, 300));
    ProduceChart(scatterChart,Scatter);
    scatterChart->setChartTitle("Scatter");

    Chart *barChart = xlsx.insertChart(17, 7, QSize(800, 300));
    ProduceChart(barChart,Bar);
    barChart->setChartTitle("Bar");

    Chart *pieChart = xlsx.insertChart(31, 7, QSize(800, 300));
    ProduceChart(pieChart,Pie);
    pieChart->setChartTitle("Pie");

    Chart *lineChart = xlsx.insertChart(45, 7, QSize(800, 300));
    ProduceChart(lineChart,Line);
    lineChart->setChartTitle("Line");

    for(int i=2;i<7;i++)
        xlsx.setColumnHidden(i,true);

    xlsx.saveAs("Test.xlsx");
    return 0;
    // return a.exec();

}

Chart* ProduceChart (Chart *obj, Type charttype)
{
    switch(charttype)
    {
        case Scatter:
        case Line:obj->setChartType(Chart::CT_ScatterChart);break;
        case Bar:obj->setChartType(Chart::CT_BarChart);break;
        case Pie:obj->setChartType(Chart::CT_PieChart);break;
    }
    //obj->setChartType(Chart::CT_ScatterChart);
    obj->setAxisTitle(obj->Left, "Data");
    obj->setAxisTitle(obj->Bottom, "X-variable");

    obj->addSeries(CellRange(1,1,50,1));
    if(charttype == Scatter || charttype == Line)
    {
        obj->addSeries(CellRange(1,2,50,2));
        obj->addSeries(CellRange(1,3,50,3));
        obj->addSeries(CellRange(1,4,50,4));
        obj->addSeries(CellRange(1,5,50,5));
        obj->addSeries(CellRange(1,6,50,6));
    }

    return obj;
}

