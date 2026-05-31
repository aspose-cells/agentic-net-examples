using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    class CombinedColumnLineChart
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Month");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["A5"].PutValue("Apr");

                sheet.Cells["B1"].PutValue("Sales");          // Column series (primary axis)
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(170);
                sheet.Cells["B5"].PutValue(200);

                sheet.Cells["C1"].PutValue("Growth %");      // Line series (secondary axis)
                sheet.Cells["C2"].PutValue(5);
                sheet.Cells["C3"].PutValue(7);
                sheet.Cells["C4"].PutValue(6);
                sheet.Cells["C5"].PutValue(8);

                // Add a chart (initially a column chart)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Add the column series (primary axis)
                chart.NSeries.Add("B2:B5", true);

                // Add the line series (will be moved to secondary axis)
                chart.NSeries.Add("C2:C5", true);

                // Set category (X) axis data
                chart.NSeries.CategoryData = "A2:A5";

                // Configure the second series as a line and plot it on the secondary Y axis
                chart.NSeries[1].PlotOnSecondAxis = true;
                chart.NSeries[1].Type = ChartType.Line;

                // Set axis titles for clarity
                chart.ValueAxis.Title.Text = "Sales";
                // If the Aspose.Cells version supports a secondary axis title, uncomment the line below:
                // chart.SecondaryValueAxis.Title.Text = "Growth %";

                // Save the workbook
                workbook.Save("CombinedColumnLineChart.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}