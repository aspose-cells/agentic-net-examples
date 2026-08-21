// Title: Add a Signature Line to an Excel Chart Sheet with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, populate data, insert a column chart, configure a SignatureLine object (signer, title, email, comments, signed date), position it below the chart, and save the file so the signature renders correctly in Excel.
// Keywords: Aspose.Cells signature line chart .NET | add signature line to Excel chart | Aspose.Cells chart shape signature | C# Aspose.Cells signature line placement | render signature line in Excel | digital signature Excel chart Aspose | Aspose.Cells chart sheet protection | US Aspose.Cells examples | EU Aspose.Cells documentation
// Common Searches: how to add a signature line to a chart using Aspose.Cells | signature line not showing on chart sheet Aspose.Cells | Aspose.Cells place signature below chart | C# code for signature line on Excel chart | verify signature line rendering in Excel with Aspose
// Developer Intent: Insert a SignatureLine shape beneath a chart in an Excel workbook and confirm it appears correctly when the file is opened.
// Use Cases: Financial dashboards that require approver signatures on each chart. | Compliance reports where charts must be signed for audit trails. | Batch generation of signed charts for multiple departments.
// AI Prompts: Generate C# code that adds a SignatureLine below a chart using Aspose.Cells and ensures proper positioning. | Explain the SignatureLine properties needed for a chart sheet and how to test its visibility in Excel. | Provide a script to programmatically verify that a signature line shape exists on a chart sheet after saving.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;          // Required for Chart and ChartType
using Aspose.Cells.Drawing;        // Required for SignatureLine

namespace AsposeCellsSignatureOnChart
{
    // Demonstrates how to create a workbook, populate data, insert a column chart, configure a SignatureLine object (signer, title, email, comments, signed date), position it below the chart, and save the file so the signature renders correctly in Excel.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart to the worksheet (rows/columns are zero‑based)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 1, 20, 10);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Create and configure a signature line
                SignatureLine signatureLine = new SignatureLine
                {
                    Signer = "John Doe",
                    Title = "Approver",
                    Email = "john.doe@example.com",
                    IsLine = true,
                    AllowComments = true,
                    ShowSignedDate = true,
                    Instructions = "Please sign to approve the chart."
                };

                // Add the signature line below the chart (row 22, column 2)
                sheet.Shapes.AddSignatureLine(22, 2, signatureLine);

                // Save the workbook
                string outputPath = "ChartWithSignatureLine.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
