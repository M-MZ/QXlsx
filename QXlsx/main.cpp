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

for (int i=2; i<51; ++i) {
xlsx.write(i, 1, rand() % 30 + 90); //A1:A9
xlsx.write(i, 2, 177); //B1:B9
xlsx.write(i, 3, 47); //C1:C9
xlsx.write(i, 4, 105); //B1:B9
xlsx.write(i, 5, 150); //C1:C9
xlsx.write(i, 6, 50); //C1:C9
}

Chart *scatterChart = xlsx.insertChart(3, 7, QSize(1000, 500));
scatterChart->setChartType(Chart::CT_ScatterChart);
scatterChart->setAxisTitle(scatterChart->Left, "Y-Variable");
scatterChart->setAxisTitle(scatterChart->Bottom, "X-variable");
scatterChart->setChartTitle("Scatter Chart");
//Will generate three lines.
scatterChart->addSeries(CellRange(1,1,50,1));
scatterChart->addSeries(CellRange(1,2,50,2));
scatterChart->addSeries(CellRange(1,3,50,3));
scatterChart->addSeries(CellRange(1,4,50,4));
scatterChart->addSeries(CellRange(1,5,50,5));
scatterChart->addSeries(CellRange(1,6,50,6));


xlsx.saveAs("Test.xlsx");
return 0;
// return a.exec();


}
