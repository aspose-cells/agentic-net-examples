using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Settings; // Namespace for ChartGlobalizationSettings

namespace AsposeCellsExamples
{
    public class ExtractOtherLabelFromPieChart
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for a pie chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["A5"].PutValue("Other"); // Category that will be rendered as "Other"

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(30);
                sheet.Cells["B3"].PutValue(45);
                sheet.Cells["B4"].PutValue(15);
                sheet.Cells["B5"].PutValue(10); // Small slice that may be grouped as "Other" depending on chart settings

                // Add a pie chart
                int chartIndex = sheet.Charts.Add(ChartType.Pie, 7, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B5", true);          // Values
                chart.NSeries.CategoryData = "A2:A5";     // Categories

                // Force chart calculation to render labels
                chart.Calculate();

                // Retrieve the default "Other" label text using globalization settings
                ChartGlobalizationSettings globalization = new ChartGlobalizationSettings();
                string otherLabel = globalization.GetOtherName();

                // Output the retrieved label
                Console.WriteLine("Default 'Other' label text: " + otherLabel);

                // Save the workbook (lifecycle: save)
                string outputPath = "OtherLabelExtraction.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExtractOtherLabelFromPieChart.Run();
        }
    }
}