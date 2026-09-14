// Title: Delete rows 30‑42, insert five blank rows, and save the worksheet as PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that removes rows 30 through 42, inserts five new rows at the same position, and exports the worksheet to a PDF file. | Generate a C# example showing how to delete a specific row range, add rows with default height, and convert the workbook to PDF using Aspose.Cells.
// Common Searches: c# aspocells delete rows 30 to 42 then insert rows before PDF export | how to remove a block of rows and add new rows in Excel with Aspose.Cells .NET | Aspose.Cells C# delete multiple rows and create blank rows prior to PDF conversion | export worksheet to PDF after row manipulation using Aspose.Cells
// Tags: delete rows range Aspose.Cells C# | insert rows default height Aspose.Cells | export worksheet to PDF Aspose.Cells | row manipulation before PDF conversion Aspose.Cells

using System;
using Aspose.Cells;

// Loads an Excel workbook, deletes rows 30‑42, inserts five new rows at that position, and saves the result as a PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Work with the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Delete rows 30‑42 (1‑based). Zero‑based start index = 29, total rows = 13
        cells.DeleteRows(29, 13);

        // Insert five new rows at the same position (row 30)
        cells.InsertRows(29, 5);

        // Export the workbook (or the specific worksheet) to PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
