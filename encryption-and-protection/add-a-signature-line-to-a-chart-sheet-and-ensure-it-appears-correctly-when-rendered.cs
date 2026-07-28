// Title: C# – Add a Signature Line to an Excel Chart Sheet with Aspose.Cells
// Description: Demonstrates how to create a workbook, populate data, generate a column chart, and place a configurable SignatureLine shape on a chart sheet using Aspose.Cells for .NET so the signature is visible when the file is opened in Excel.
// Keywords: Aspose.Cells signature line C# | add signature line to chart sheet | Excel chart sheet digital signature | Aspose.Cells shape example | .NET Excel signature line | chart sheet signature Aspose | C# Excel compliance signature
// Common Searches: how to insert a signature line on an Excel chart sheet using Aspose.Cells | Aspose.Cells C# add signature line to specific cell | render signature line on chart sheet .NET | Aspose.Cells place digital signature on chart | C# code for signature line in Excel workbook
// Developer Intent: Insert a customizable SignatureLine shape onto a chart sheet and ensure it renders correctly in Excel.
// Use Cases: Create a financial dashboard chart that includes an approver’s signature for audit trails. | Automate compliance reports where a compliance officer signs off directly on the chart sheet. | Generate sales performance charts with a certified signature line before distribution.
// AI Prompts: Write C# code with Aspose.Cells to add a signature line at cell D30 on a chart sheet and save as .xlsx. | Explain how to configure SignatureLine properties (Signer, Title, Email, ShowSignedDate, etc.) for proper display on a chart sheet. | Provide steps to verify that the signature line appears correctly in Excel after saving the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Charts;

namespace AsposeCellsSignatureOnChartSheet
{
    // Demonstrates how to create a workbook, populate data, generate a column chart, and place a configurable SignatureLine shape on a chart sheet using Aspose.Cells for .NET so the signature is visible when the file is opened in Excel.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet (will act as a chart sheet)
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 1, 20, 10);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Create a signature line object and set its properties
                SignatureLine signatureLine = new SignatureLine
                {
                    Signer = "John Doe",
                    Title = "Approver",
                    Email = "john.doe@example.com",
                    IsLine = true,
                    AllowComments = true,
                    ShowSignedDate = true,
                    Instructions = "Please sign to confirm the chart."
                };

                // Add the signature line to the worksheet (chart sheet) at a specific cell location
                // Here we place it at row 22, column 2 (cell B23) – adjust as needed
                sheet.Shapes.AddSignatureLine(22, 2, signatureLine);

                // Save the workbook; the signature line will be visible when the file is opened in Excel
                workbook.Save("ChartSheetWithSignatureLine.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
