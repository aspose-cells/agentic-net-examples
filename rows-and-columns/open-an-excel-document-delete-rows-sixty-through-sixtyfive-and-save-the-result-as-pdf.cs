// Title: Delete rows 60‑65 from an Excel worksheet and export to PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Create a C# program with Aspose.Cells that opens an .xlsx file, removes the rows numbered 60 through 65 from the first sheet, and saves the workbook as a PDF. | Show how to use Aspose.Cells in C# to programmatically delete a block of rows and then generate a PDF document from the modified worksheet. | Write Aspose.Cells code that loads a workbook, eliminates a specific row range, and outputs the result as a PDF file in C#.
// Common Searches: Aspose.Cells C# delete rows 60‑65 before converting to PDF | how to remove a range of rows in Excel with Aspose.Cells and save as PDF | C# example for deleting specific rows in a worksheet and exporting to PDF using Aspose.Cells | Aspose.Cells delete rows 60 to 65 and generate PDF output | remove rows from Excel file programmatically and convert to PDF in .NET
// Tags: Aspose.Cells delete rows C# | Excel row removal PDF conversion Aspose.Cells | C# delete specific row range Aspose.Cells | Aspose.Cells export worksheet to PDF after row deletion | remove rows 60-65 Aspose.Cells

using Aspose.Cells;

// // Loads input.xlsx, deletes rows 60‑65 from the first worksheet, and saves the modified workbook as output.pdf using Aspose.Cells.
class Program
{
    static void Main()
    {
        // Load the existing Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Delete rows 60 through 65 (zero‑based index: start at 59, delete 6 rows)
        sheet.Cells.DeleteRows(59, 6);

        // Save the modified workbook as a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
