// Title: Remove all empty rows from an Excel worksheet and save the result as a PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Load an .xlsx file, invoke DeleteBlankRows on the desired worksheet, then export the workbook to PDF with Aspose.Cells in C#. | Using Aspose.Cells in a .NET project, compact a worksheet by deleting rows without data and generate a PDF output.
// Common Searches: C# Aspose.Cells delete rows that have no data and export to PDF | How to remove blank rows from an Excel sheet before PDF conversion using Aspose.Cells | Aspose.Cells DeleteBlankRows method example for .NET | Convert cleaned Excel workbook to PDF with Aspose.Cells C#
// Tags: DeleteBlankRows Aspose.Cells C# | Excel to PDF conversion Aspose.Cells | remove empty rows worksheet Aspose.Cells | compact workbook before PDF export .NET | SaveFormat.Pdf usage Aspose.Cells | worksheet row cleanup Aspose.Cells

using System;
using Aspose.Cells;

// // Loads an Excel file, removes all completely empty rows from the first worksheet, and saves the cleaned workbook as a PDF.
class Program
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Get the first worksheet (or any specific worksheet you need)
        Worksheet worksheet = workbook.Worksheets[0];

        // Delete all blank rows that contain no data or objects
        worksheet.Cells.DeleteBlankRows();

        // Save the compacted workbook as a PDF file
        string outputPath = "output.pdf";
        workbook.Save(outputPath, SaveFormat.Pdf);
    }
}
