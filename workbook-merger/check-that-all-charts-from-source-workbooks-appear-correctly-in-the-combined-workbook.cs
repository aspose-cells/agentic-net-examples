using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartCombineCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create first source workbook with a chart
            Workbook sourceWorkbook1 = new Workbook();
            Worksheet sheet1 = sourceWorkbook1.Worksheets[0];
            sheet1.Name = "Source1";
            // Populate data
            sheet1.Cells["A1"].PutValue("Category");
            sheet1.Cells["A2"].PutValue("A");
            sheet1.Cells["A3"].PutValue("B");
            sheet1.Cells["A4"].PutValue("C");
            sheet1.Cells["B1"].PutValue("Value");
            sheet1.Cells["B2"].PutValue(10);
            sheet1.Cells["B3"].PutValue(20);
            sheet1.Cells["B4"].PutValue(30);
            // Add a chart
            int chartIndex1 = sheet1.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart1 = sheet1.Charts[chartIndex1];
            chart1.NSeries.Add("B2:B4", true);
            chart1.NSeries.CategoryData = "A2:A4";
            chart1.Title.Text = "Source1 Chart";

            // Create second source workbook with a chart
            Workbook sourceWorkbook2 = new Workbook();
            Worksheet sheet2 = sourceWorkbook2.Worksheets[0];
            sheet2.Name = "Source2";
            // Populate data
            sheet2.Cells["A1"].PutValue("Month");
            sheet2.Cells["A2"].PutValue("Jan");
            sheet2.Cells["A3"].PutValue("Feb");
            sheet2.Cells["A4"].PutValue("Mar");
            sheet2.Cells["B1"].PutValue("Sales");
            sheet2.Cells["B2"].PutValue(150);
            sheet2.Cells["B3"].PutValue(200);
            sheet2.Cells["B4"].PutValue(250);
            // Add a chart
            int chartIndex2 = sheet2.Charts.Add(ChartType.Line, 5, 0, 20, 8);
            Chart chart2 = sheet2.Charts[chartIndex2];
            chart2.NSeries.Add("B2:B4", true);
            chart2.NSeries.CategoryData = "A2:A4";
            chart2.Title.Text = "Source2 Chart";

            // Destination workbook that will receive the combined content
            Workbook destWorkbook = new Workbook();

            // Combine both source workbooks into the destination workbook
            destWorkbook.Combine(sourceWorkbook1);
            destWorkbook.Combine(sourceWorkbook2);

            // Verify that all charts from the sources are present in the combined workbook
            int expectedChartCount = sourceWorkbook1.Worksheets[0].Charts.Count + sourceWorkbook2.Worksheets[0].Charts.Count;
            int actualChartCount = 0;

            // Iterate through all worksheets in the combined workbook and sum chart counts
            foreach (Worksheet ws in destWorkbook.Worksheets)
            {
                actualChartCount += ws.Charts.Count;
            }

            Console.WriteLine($"Expected total charts: {expectedChartCount}");
            Console.WriteLine($"Actual total charts in combined workbook: {actualChartCount}");

            // Simple validation: compare chart types and titles
            bool allChartsMatch = true;
            int chartCounter = 0;

            // Helper to collect source charts in order
            Chart[] sourceCharts = new Chart[]
            {
                chart1,
                chart2
            };

            foreach (Worksheet ws in destWorkbook.Worksheets)
            {
                foreach (Chart combinedChart in ws.Charts)
                {
                    if (chartCounter >= sourceCharts.Length)
                    {
                        allChartsMatch = false;
                        break;
                    }

                    Chart srcChart = sourceCharts[chartCounter];

                    // Compare type
                    if (combinedChart.Type != srcChart.Type)
                    {
                        allChartsMatch = false;
                    }

                    // Compare title text
                    if (combinedChart.Title.Text != srcChart.Title.Text)
                    {
                        allChartsMatch = false;
                    }

                    chartCounter++;
                }
            }

            Console.WriteLine($"All charts correctly transferred: {allChartsMatch}");

            // Save the combined workbook for visual inspection (optional)
            destWorkbook.Save("CombinedWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}