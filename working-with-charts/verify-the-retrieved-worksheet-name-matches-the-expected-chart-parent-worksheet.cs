// Title: C# – Verify a Chart’s Parent Worksheet Name with Aspose.Cells for .NET
// Description: Creates a workbook, renames the first worksheet, adds a column chart, retrieves the chart’s parent worksheet via Chart.Worksheet, compares the name to the expected value, outputs the result, and saves the file.
// Keywords: Aspose.Cells chart worksheet name | Chart.Worksheet property C# | verify chart parent worksheet | Aspose.Cells .NET chart verification | retrieve chart parent sheet
// Common Searches: Aspose.Cells get chart worksheet name C# | How to check chart parent sheet in Aspose.Cells | Validate chart worksheet name Aspose.Cells .NET | Chart.Worksheet.Name example Aspose.Cells | Verify chart location before saving workbook
// Developer Intent: Confirm that a chart’s Worksheet.Name matches a predefined worksheet name.
// Use Cases: Automated tests to ensure charts are placed on the correct sheet. | Pre‑publish validation of report workbooks that contain charts. | Conditional formatting or data binding that depends on a specific chart sheet.
// AI Prompts: Generate C# code using Aspose.Cells to read a chart’s parent worksheet name and compare it with an expected string. | Explain how to log or throw an exception when a chart’s worksheet name does not match the expected value. | Show a loop that iterates through all charts in a workbook and validates each chart’s Worksheet.Name against a naming convention.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartWorksheetVerification
{
    // Creates a workbook, renames the first worksheet, adds a column chart, retrieves the chart’s parent worksheet via Chart.Worksheet, compares the name to the expected value, outputs the result, and saves the file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet and set a known name
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "ExpectedSheet";

            // Add a sample chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = worksheet.Charts[chartIndex];

            // Retrieve the worksheet name via the chart's Worksheet property
            string chartParentWorksheetName = chart.Worksheet.Name;

            // Verify that the retrieved name matches the expected worksheet name
            if (chartParentWorksheetName == worksheet.Name)
            {
                Console.WriteLine("Verification successful: Chart's parent worksheet name matches expected name.");
            }
            else
            {
                Console.WriteLine($"Verification failed: Expected '{worksheet.Name}' but got '{chartParentWorksheetName}'.");
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save("ChartWorksheetVerification_out.xlsx");
        }
    }
}
