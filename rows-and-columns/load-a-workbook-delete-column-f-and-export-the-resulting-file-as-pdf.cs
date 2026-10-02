// Title: Delete column F from an Excel worksheet and save the workbook as a PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, removes column F from the first sheet, and saves the result as a PDF. | Generate a snippet that uses Aspose.Cells to delete a zero‑based column index (5) from a worksheet and then export the workbook to PDF. | Provide an example of loading a workbook, calling Cells.DeleteColumn on column 5, and calling Workbook.Save with SaveFormat.Pdf in C#.
// Common Searches: Aspose.Cells C# delete column F from Excel and export to PDF | How to remove a specific column from an Excel file and save as PDF using Aspose.Cells | C# code to delete column index 5 in a worksheet and save workbook as PDF with Aspose.Cells | Aspose.Cells example for column deletion and PDF conversion in .NET
// Tags: Aspose.Cells column removal C# | Aspose.Cells PDF export after column removal | remove specific column from Excel worksheet | Cells.DeleteColumn usage example | save workbook as PDF Aspose.Cells

using Aspose.Cells;

// The program loads 'input.xlsx' into an Aspose.Cells Workbook, deletes column F (zero‑based index 5) from the first worksheet, and saves the modified workbook as 'output.pdf' using the PDF save format.
class Program
{
    static void Main()
    {
        // Load the workbook from the source file
        Workbook workbook = new Workbook("input.xlsx");

        // Delete column F (zero‑based index 5) from the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells.DeleteColumn(5);

        // Export the modified workbook as a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
