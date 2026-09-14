// Title: Delete column Z, hide rows 50‑55, and export Excel to PDF using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an .xlsx workbook with Aspose.Cells, removes column Z, conceals rows 50‑55, and saves the file as a PDF. | Write a .NET snippet using Aspose.Cells to delete a specific column, hide a range of rows, and export the worksheet to PDF. | Create a C# program that loads a spreadsheet, calls Cells.DeleteColumn for column Z, calls Cells.HideRows for rows 50‑55, then saves the result as a PDF document.
// Common Searches: Aspose.Cells C# how to delete a column and hide rows before converting to PDF | remove column Z and hide rows 50 to 55 in Excel using Aspose.Cells .NET | export modified Excel sheet to PDF after column removal and row hiding with Aspose.Cells
// Tags: delete column Z Aspose.Cells C# | cells.hiderows for rows 50-55 Aspose.Cells | pdf conversion after worksheet modification Aspose.Cells | excel column removal and row concealment .NET | cells.deletecolumn usage example

using System;
using Aspose.Cells;

// Loads input.xlsx, deletes column Z, hides rows 50‑55, and saves the workbook as output.pdf using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Paths to the input Excel file and the output PDF file
        string inputPath = "input.xlsx";
        string outputPath = "output.pdf";

        // Load the existing workbook
        Workbook workbook = new Workbook(inputPath);

        // Work with the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Delete column Z (zero‑based index 25, where A = 0)
        cells.DeleteColumn(25);

        // Hide rows 50 through 55 (zero‑based start index 49, total 6 rows)
        cells.HideRows(49, 6);

        // Save the modified workbook as a PDF document
        workbook.Save(outputPath, SaveFormat.Pdf);
    }
}
