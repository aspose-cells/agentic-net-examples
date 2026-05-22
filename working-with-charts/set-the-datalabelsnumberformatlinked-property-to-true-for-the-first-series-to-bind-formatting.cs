using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class SetNumberFormatLinkedDemo
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(150);
            sheet.Cells["B3"].PutValue(250);

            // Cells with formatted text that will be linked
            sheet.Cells["C1"].PutValue("Formatted");
            sheet.Cells["C2"].PutValue("150 units");
            sheet.Cells["C3"].PutValue("250 units");

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Add the first series and set its category data
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Access the first series
            Series firstSeries = chart.NSeries[0];

            // Enable data labels and link them to the formatted cells
            firstSeries.DataLabels.ShowValue = true;
            firstSeries.DataLabels.LinkedSource = "C2:C3";

            // Bind the number format of the data labels to the source cells
            firstSeries.DataLabels.NumberFormatLinked = true;

            // Define output file path
            string outputPath = "SetNumberFormatLinkedDemo.xlsx";

            // Save the workbook (overwrite if exists)
            workbook.Save(outputPath);
        }
    }
}