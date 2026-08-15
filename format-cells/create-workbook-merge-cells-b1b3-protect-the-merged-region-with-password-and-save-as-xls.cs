// Title: Aspose.Cells for .NET – Create XLS Workbook, Merge B1:B3, and Apply Password Protection
// Description: C# sample that creates a new workbook, merges the range B1:B3 on the first worksheet, protects the merged area with a password, and saves the result as an XLS file.
// Keywords: Aspose.Cells | C# merge cells | B1:B3 merge | worksheet protection | password protect | save as XLS | Aspose.Cells .NET example | protect merged region | Excel file generation | Aspose.Cells tutorial
// Common Searches: Aspose.Cells merge cells B1:B3 C# | protect merged cells with password Aspose.Cells | save protected workbook as .xls using Aspose.Cells .NET | C# lock merged range in Excel Aspose.Cells | worksheet protection password example Aspose.Cells
// Developer Intent: Create an XLS workbook, merge cells B1:B3, and secure the merged range with a password using Aspose.Cells for .NET.
// Use Cases: Distribute a template where the title row (B1:B3) is merged and locked to preserve layout. | Generate a financial report with a read‑only merged header while allowing data entry elsewhere. | Provide external partners an XLS file where specific merged sections are protected to prevent accidental changes.
// AI Prompts: Show C# code that merges B1:B3 and applies password protection with Aspose.Cells for .NET. | How can I lock only a merged region in an Excel file while keeping other cells editable using Aspose.Cells? | Explain the steps to save a password‑protected workbook as an XLS file and ensure the protection persists.

using System;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // C# sample that creates a new workbook, merges the range B1:B3 on the first worksheet, protects the merged area with a password, and saves the result as an XLS file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Merge cells B1:B3 (zero‑based indices: row 0, column 1, 3 rows, 1 column)
                sheet.Cells.Merge(0, 1, 3, 1);

                // Protect the worksheet (including the merged region) with a password
                sheet.Protect(ProtectionType.All, "SecretPassword", null);

                // Save the workbook; format is inferred from the file extension
                string outputPath = "MergedProtected.xls";
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
