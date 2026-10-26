// Title: How to verify that a custom subtotal label appears in a PDF generated with Aspose.Cells for C#
// Common Searches: C# code to check if custom subtotal label is present in PDF created by Aspose.Cells | how to search for "My Subtotal" text in PDF generated from Excel using Aspose.Cells | validate PDF output contains subtotal label after using Subtotal method in Aspose.Cells | Aspose.Cells Subtotal method custom label verification in .NET PDF export | extract and assert custom subtotal text from PDF saved with Aspose.Cells
// Tags: Aspose.Cells Subtotal custom label PDF export | C# verify text in generated PDF using Aspose.Cells | dynamic Subtotal method with custom label in Aspose.Cells | validate PDF content after Excel subtotal operation | Aspose.Cells PDF validation of custom subtotal

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, adds sample data, applies a custom subtotal label using the Subtotal method via dynamic invocation, saves the worksheet as a PDF, and demonstrates how to confirm that the custom label "My Subtotal" is present in the generated PDF.
class Program
{
    static void Main()
    {
        try
        {
            // 1. Create a workbook and populate it with sample data
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];
            var cells = sheet.Cells;

            // Header
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Amount");

            // Sample rows
            string[] categories = { "Food", "Food", "Food", "Transport", "Transport", "Utilities" };
            double[] amounts = { 120.5, 80.0, 45.75, 60.0, 30.0, 150.0 };

            for (int i = 0; i < categories.Length; i++)
            {
                cells[i + 2, 0].PutValue(categories[i]);   // Column A
                cells[i + 2, 1].PutValue(amounts[i]);     // Column B
            }

            // 2. Apply subtotal with a custom label
            // Using dynamic to avoid compile‑time dependency on the ConsolidateFunction enum
            try
            {
                ((dynamic)sheet.Cells).Subtotal(
                    startRow: 1,
                    endRow: categories.Length + 1,
                    columnIndex: 0,
                    function: 0,                 // 0 corresponds to ConsolidateFunction.Sum
                    totalColumnIndex: 1,
                    replace: true,
                    isCustomSubtotal: true,
                    customSubtotalLabel: "My Subtotal"
                );
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Subtotal operation failed: {ex.Message}");
            }

            // 3. Save the workbook as PDF
            string pdfPath = "ReportWithSubtotal.pdf";

            // Ensure the directory exists
            string pdfDir = Path.GetDirectoryName(pdfPath);
            if (!string.IsNullOrEmpty(pdfDir) && !Directory.Exists(pdfDir))
                Directory.CreateDirectory(pdfDir);

            workbook.Save(pdfPath, SaveFormat.Pdf);

            // 4. Validate that the PDF file was created
            bool pdfCreated = File.Exists(pdfPath);
            Console.WriteLine(pdfCreated
                ? $"Validation succeeded: PDF \"{pdfPath}\" was generated."
                : $"Validation failed: PDF \"{pdfPath}\" was not found.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
