// Title: Hide columns 5‑8, unhide columns 6‑7 with a custom width, and export the worksheet to PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that loads an Excel workbook, hides columns 5 through 8, then unhides columns 6 and 7 setting their width to 15 characters, and finally saves the sheet as a PDF. | Generate a .NET example showing how to manipulate column visibility (hide a range, unhide a sub‑range with specific width) and convert the worksheet to PDF using Aspose.Cells.
// Common Searches: Aspose.Cells C# hide columns 5 to 8 then unhide 6 and 7 with width 15 before exporting to PDF | How to set column width while unhiding specific columns using Aspose.Cells in .NET | Export an Excel worksheet to PDF after adjusting column visibility with Aspose.Cells C#
// Tags: hide columns range Aspose.Cells C# | unhide columns with specific width Aspose.Cells | export worksheet to PDF Aspose.Cells | column visibility manipulation Aspose.Cells | set column width Aspose.Cells C#

using Aspose.Cells;

// // Load input.xlsx, hide columns 5‑8, unhide columns 6‑7 with a width of 15 characters, and save the worksheet as output.pdf.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Hide columns 5 through 8 (1‑based). Zero‑based start index = 4, total = 4.
        cells.HideColumns(4, 4);

        // Unhide columns 6 and 7 (1‑based) with a width of 15 characters.
        cells.UnhideColumns(5, 2, 15.0);

        // Export the worksheet to PDF.
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
