using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartAutomation
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add multiple worksheets (for demonstration)
            Worksheet sheet1 = workbook.Worksheets[0]; // default sheet
            sheet1.Name = "Sheet1";

            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");

            // Populate each worksheet with sample data
            PopulateSampleData(sheet1);
            PopulateSampleData(sheet2);
            PopulateSampleData(sheet3);

            // Loop through all worksheets and add a column chart to each
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Define the position of the chart (topRow, leftColumn, bottomRow, rightColumn)
                int topRow = 5;
                int leftColumn = 2;
                int bottomRow = 25;
                int rightColumn = 11;

                // Add a column chart to the current worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, topRow, leftColumn, bottomRow, rightColumn);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart (assumes data is in A1:B5)
                // First series values
                chart.NSeries.Add("=Sheet" + (sheet.Index + 1) + "!$B$2:$B$5", true);
                // Category (X) axis data
                chart.NSeries.CategoryData = "=Sheet" + (sheet.Index + 1) + "!$A$2:$A$5";

                // Optional: give the chart a title
                chart.Title.Text = $"Sample Chart for {sheet.Name}";

                // Recalculate the chart layout
                chart.Calculate();
            }

            // Save the workbook to a file
            workbook.Save("ChartsForAllSheets.xlsx", SaveFormat.Xlsx);
        }

        // Helper method to fill a worksheet with simple sample data
        private static void PopulateSampleData(Worksheet sheet)
        {
            // Header
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");

            // Sample rows
            for (int i = 2; i <= 5; i++)
            {
                sheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
                sheet.Cells[$"B{i}"].PutValue((i - 1) * 10); // 10, 20, 30, 40
            }
        }
    }
}