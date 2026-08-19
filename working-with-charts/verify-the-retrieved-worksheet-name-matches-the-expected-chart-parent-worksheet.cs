// Title: Validate a Chart’s Parent Worksheet Name with Aspose.Cells for .NET
// Description: Creates a workbook, renames the first sheet, adds a column chart, retrieves the chart’s parent worksheet via the Chart.Worksheet property, compares it to an expected name, and prints a verification message. Shows how to confirm chart‑worksheet association in C#.
// Keywords: Aspose.Cells | Chart.Worksheet | parent worksheet name | C# chart verification | .NET workbook chart | retrieve chart worksheet | compare chart sheet name | chart placement validation
// Common Searches: Aspose.Cells get chart worksheet name C# | How to verify chart parent sheet in .NET | Chart.Worksheet property example | Check chart belongs to specific worksheet Aspose.Cells | Validate chart location after copying worksheets
// Developer Intent: Confirm that a chart’s parent worksheet name matches a predefined value.
// Use Cases: Automated tests that ensure generated charts are placed on the correct sheet. | Validation after moving or copying charts between worksheets in multi‑sheet reports. | Runtime checks before saving a workbook to guarantee chart data references the intended sheet.
// AI Prompts: Generate C# code using Aspose.Cells that throws an exception when Chart.Worksheet.Name does not equal a given string. | Create a unit test that asserts Chart.Worksheet.Name matches the expected worksheet name for a newly added chart. | Provide a sample that logs a mismatch between a chart’s parent worksheet and the expected name, then renames the worksheet accordingly.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace VerifyChartParentWorksheet
{
    // Creates a workbook, renames the first sheet, adds a column chart, retrieves the chart’s parent worksheet via the Chart.Worksheet property, compares it to an expected name, and prints a verification message. Shows how to confirm chart‑worksheet association in C#.
    class Program
    {
        static void Main()
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

            // Expected parent worksheet name
            string expectedWorksheetName = "DataSheet";

            // Retrieve the actual parent worksheet name via Chart.Worksheet property
            string actualWorksheetName = chart.Worksheet.Name;

            // Verify that the retrieved name matches the expected name
            if (actualWorksheetName == expectedWorksheetName)
            {
                Console.WriteLine("Verification succeeded: Chart's parent worksheet name matches expected value.");
            }
            else
            {
                Console.WriteLine($"Verification failed: Expected '{expectedWorksheetName}', but got '{actualWorksheetName}'.");
            }

            // Save the workbook (optional, just to demonstrate lifecycle usage)
            workbook.Save("VerifyChartParentWorksheet_out.xlsx");
        }
    }
}
