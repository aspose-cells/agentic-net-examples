// Title: Insert a row at index 15 in an Excel worksheet and export the workbook to PDF with Aspose.Cells for .NET
// Common Searches: C# code to insert a row at a specific index in an Excel file using Aspose.Cells | How to save an edited Excel workbook as a PDF with Aspose.Cells .NET | Aspose.Cells insert rows before saving to PDF example | Convert modified Excel worksheet to PDF after adding rows in C# | Insert row at row 16 (zero based 15) with Aspose.Cells and export to PDF
// Tags: insert rows Aspose.Cells C# | excel to pdf conversion Aspose.Cells | modify worksheet row index Aspose.Cells | save workbook as PDF .NET | Aspose.Cells row insertion example

using System;
using Aspose.Cells;

// The sample loads an Excel file (input.xlsx) with Aspose.Cells, inserts a new row at zero‑based index 15 (the 16th row) on the first worksheet, and then saves the updated workbook as a PDF file (output.pdf).
class Program
{
    static void Main()
    {
        // Path to the source Excel file
        string inputPath = "input.xlsx";

        // Load the workbook
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet (you can change the index or name as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Insert a new row at index 15 (zero‑based, so this is the 16th row in Excel)
        sheet.Cells.InsertRows(15, 1);

        // Path for the resulting PDF file
        string outputPath = "output.pdf";

        // Save the workbook as PDF
        workbook.Save(outputPath, SaveFormat.Pdf);
    }
}
