#include <QCoreApplication>

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
xlsx.write(1, 1, "Data");
xlsx.write(1, 2, "B");
xlsx.write(1, 3, "C");
for (int i=2; i<11; ++i) {
xlsx.write(i, 1, i*i*i); //A1:A9
xlsx.write(i, 2, i*i); //B1:B9
xlsx.write(i, 3, i*i-1); //C1:C9
}

Chart *scatterChart = xlsx.insertChart(3, 7, QSize(300, 300));
scatterChart->setChartType(Chart::CT_ScatterChart);
scatterChart->setAxisTitle(scatterChart->Left, "MMZ");
scatterChart->setAxisTitle(scatterChart->Bottom, "X-variable");
scatterChart->setChartTitle("Scatter Chart");
//Will generate three lines.
scatterChart->addSeries(CellRange(1,1,10,1));
scatterChart->addSeries(CellRange(1,2,10,2));
scatterChart->addSeries(CellRange(1,3,10,3));


Chart *barChart = xlsx.insertChart(3,3, QSize(300,300));
barChart->setChartType(Chart::CT_BarChart);
barChart->setAxisTitle(barChart->Left, "MMZ");
barChart->setAxisTitle(barChart->Bottom, "X-variable");
barChart->setChartTitle("Bar Chart");
barChart->addSeries(CellRange(1,1,10,1));
barChart->addSeries(CellRange(1,2,10,2));
barChart->addSeries(CellRange(1,3,10,3));


Chart *pieChart = xlsx.insertChart(3, 3, QSize(300, 300));
pieChart->setChartType(Chart::CT_PieChart);
pieChart->addSeries(CellRange(1,1,9,1));
//pieChart->addSeries(CellRange(1,2,9,2));
//pieChart->addSeries(CellRange(1,3,9,3));
pieChart->setChartTitle("Pie Chart");

Chart *lineChart = xlsx.insertChart(3, 3, QSize(300, 300));
lineChart->setChartType(Chart::CT_LineChart);
lineChart->addSeries(CellRange(1,1,9,1));
lineChart->addSeries(CellRange(1,2,9,2));
lineChart->addSeries(CellRange(1,3,9,3));
lineChart->setChartTitle("Line Chart");


xlsx.saveAs("Test.xlsx");
return 0;
// return a.exec();
}
