// Title: Check Chart's Parent Worksheet Name Using Aspose.Cells for .NET (C#)
// Description: This C# example creates a workbook, renames the first sheet to "DataSheet", adds a column chart, accesses the chart's parent sheet via the Chart.Worksheet property, compares the sheet name to an expected value, logs the outcome, and saves the file.
// Keywords: aspose.cells chart worksheet | chart.workbook parent sheet c# | verify chart sheet name | aspose.cells chart worksheet property | c# chart parent worksheet validation
// Common Searches: Aspose.Cells get chart's worksheet | C# verify chart parent sheet name | Chart.Worksheet property example | how to confirm chart is on correct sheet Aspose.Cells | validate chart location in generated workbook
// Developer Intent: Determine whether the worksheet returned by Chart.Worksheet matches a predefined name.
// Use Cases: Automated tests to ensure charts are generated on the intended sheet | Quality checks after copying or moving worksheets that contain charts | Dynamic report generation where chart placement must be verified
// AI Prompts: Write C# code with Aspose.Cells that creates a workbook, adds a chart, retrieves its parent worksheet, and asserts the worksheet name equals a specified string. | Show how to loop through all charts in a workbook and validate each chart's parent sheet name using Aspose.Cells for .NET. | Provide a logging pattern for mismatched chart worksheet names, including sheet index and chart ID, in a C# Aspose.Cells application.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This C# example creates a workbook, renames the first sheet to "DataSheet", adds a column chart, accesses the chart's parent sheet via the Chart.Worksheet property, compares the sheet name to an expected value, logs the outcome, and saves the file.
    public class VerifyChartParentWorksheet
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet and give it a known name
                Worksheet worksheet = workbook.Worksheets[0];
                worksheet.Name = "DataSheet";

                // Add sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(30);

                // Add a chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = worksheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Retrieve the worksheet that contains the chart
                Worksheet chartParent = chart.Worksheet;

                // Verify that the retrieved worksheet name matches the expected name
                string expectedName = "DataSheet";
                if (chartParent.Name == expectedName)
                {
                    Console.WriteLine($"Success: Chart's parent worksheet name '{chartParent.Name}' matches expected '{expectedName}'.");
                }
                else
                {
                    Console.WriteLine($"Failure: Chart's parent worksheet name '{chartParent.Name}' does not match expected '{expectedName}'.");
                }

                // Save the workbook
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "VerifyChartParentWorksheet_out.xlsx");
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            VerifyChartParentWorksheet.Run();
        }
    }
}
