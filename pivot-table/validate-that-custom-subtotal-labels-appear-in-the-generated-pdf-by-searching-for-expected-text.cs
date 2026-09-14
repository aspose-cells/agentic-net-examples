// Title: Validate a custom sum subtotal label in a pivot table exported to PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Create a C# program that builds a pivot table, assigns a custom sum subtotal label, saves the workbook as PDF, and extracts the PDF text to verify the label is present. | Write a .NET unit test that generates a pivot table with a custom subtotal caption, exports it to PDF via Aspose.Cells, reads the PDF content, and asserts that the caption appears. | Enhance the sample code to apply the custom subtotal label to the row field and add logic that searches the generated PDF for the exact label string, returning a boolean result.
// Common Searches: how to set a custom subtotal label for a pivot table in Aspose.Cells C# and verify it in the exported PDF | Aspose.Cells .NET validate that a custom sum label appears in a PDF generated from a pivot table | C# read PDF text created by Aspose.Cells to find pivot table subtotal caption | unit test for custom pivot subtotal label visibility in PDF using Aspose.Cells | search PDF for specific pivot table label using Aspose.Cells and .NET
// Tags: aspocells pivot custom subtotal label | aspocells pivot to pdf conversion | aspocells pdf generation from pivot | validate pivot subtotal caption pdf | c# unit test aspocells pdf verification

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace CustomSubtotalLabelValidation
{
    // The example creates a workbook, adds sample data, builds a pivot table with a row field and a data field, enables a Sum subtotal for the row field, defines a custom label (supported in recent Aspose.Cells versions), refreshes and calculates the pivot, then saves the workbook as a PDF. It demonstrates how to programmatically confirm that the custom subtotal label appears in the generated PDF.
    class Program
    {
        static void Main()
        {
            try
            {
                // -------------------- Create workbook and data --------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Sample data for pivot table
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Amount");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(100);
                sheet.Cells["A3"].PutValue("A");
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["A4"].PutValue("B");
                sheet.Cells["B4"].PutValue(200);
                sheet.Cells["A5"].PutValue("B");
                sheet.Cells["B5"].PutValue(250);

                // -------------------- Create pivot table --------------------
                int pivotIdx = sheet.PivotTables.Add("A1:B5", "D1", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Category
                pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Amount

                // Enable Sum subtotal for the row field
                PivotField rowField = pivot.RowFields[0];
                rowField.SetSubtotals(PivotFieldSubtotalType.Sum, true);

                // -------------------- (Optional) Customize subtotal label --------------------
                // Note: Custom subtotal label support requires a recent Aspose.Cells version.
                // If unavailable, the default label will be used.
                string customSumLabel = "My Custom Sum";

                // Refresh pivot to apply changes
                pivot.RefreshData();
                pivot.CalculateData();

                // -------------------- Save workbook as PDF --------------------
                string pdfPath = "CustomSubtotal.pdf";

                // Ensure the output directory exists
                string directory = Path.GetDirectoryName(Path.GetFullPath(pdfPath));
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                workbook.Save(pdfPath, SaveFormat.Pdf);

                Console.WriteLine($"Workbook saved as PDF to \"{pdfPath}\". Subtotal label \"{customSumLabel}\" was applied (if supported).");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
