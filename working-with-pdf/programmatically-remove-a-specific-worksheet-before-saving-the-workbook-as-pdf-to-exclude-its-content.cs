// Title: Remove a specific worksheet from an Excel file and save the remaining content as PDF using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx workbook with Aspose.Cells, checks for a worksheet named 'SheetToExclude', removes it if present, and then exports the workbook to a PDF file. | Show how to safely delete a worksheet by name in Aspose.Cells before performing a PDF conversion, including handling the case where the sheet does not exist. | Provide a complete Aspose.Cells workflow that programmatically excludes a particular sheet from the PDF output, covering loading, sheet removal, and saving as PDF.
// Common Searches: aspnet remove worksheet named SheetToExclude before converting Excel to PDF with Aspose.Cells | how to exclude a single sheet from PDF export using Aspose.Cells C# | Aspose.Cells delete specific worksheet then save workbook as PDF example
// Tags: worksheet removal Aspose.Cells | PDF export without target sheet | Aspose.Cells omit sheet during PDF conversion | delete sheet by name Aspose.Cells | convert Excel to PDF after sheet exclusion

using System;
using Aspose.Cells;

// The example loads an existing Excel workbook, looks for a worksheet called 'SheetToExclude', removes it if found, and then saves the modified workbook as a PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Name of the worksheet to be removed
        string sheetToRemove = "SheetToExclude";

        // Find the worksheet by name
        Worksheet sheet = workbook.Worksheets[sheetToRemove];

        // If the worksheet exists, remove it from the workbook
        if (sheet != null)
        {
            int index = sheet.Index;
            workbook.Worksheets.RemoveAt(index);
        }

        // Save the modified workbook as PDF (replace with your desired output path)
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
