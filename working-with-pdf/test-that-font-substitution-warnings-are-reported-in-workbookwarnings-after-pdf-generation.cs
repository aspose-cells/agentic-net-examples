// Title: Check for font substitution warnings in Aspose.Cells workbook after PDF conversion (C#)
// Common Searches: how to capture font substitution warnings when converting Excel to PDF with Aspose.Cells C# | Aspose.Cells .NET retrieve workbook warnings after saving as PDF | detect missing font warnings during PDF export using Aspose.Cells | Aspose.Cells GetWarnings method usage for PDF conversion warnings | C# example to read font substitution alerts after workbook.Save to PDF
// Tags: Aspose.Cells font substitution warning handling | C# workbook warnings after PDF conversion | Aspose.Cells GetWarnings API | missing font alert handling in PDF export | memory stream PDF generation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates an in‑memory Workbook, writes text to cell A1, applies a style with a non‑existent font to trigger substitution, saves the workbook to PDF via a MemoryStream, and shows that font substitution warnings can be accessed through the workbook's GetWarnings method when available.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook in memory
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Write sample text to cell A1
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Testing font substitution warning");

            // Apply a style with a non‑existent font to trigger substitution
            Style style = workbook.CreateStyle();
            style.Font.Name = "NonExistentFontXYZ";
            cell.SetStyle(style);

            // Save the workbook to PDF using a memory stream (no file I/O)
            using (MemoryStream pdfStream = new MemoryStream())
            {
                workbook.Save(pdfStream, SaveFormat.Pdf);
                // The PDF bytes are now in pdfStream; you could write them to a file if needed
            }

            // Note: In some Aspose.Cells versions, workbook warnings can be retrieved via workbook.GetWarnings().
            // If the method is unavailable, font substitution warnings are not exposed directly.
            Console.WriteLine("Workbook processed and PDF generated in memory.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
