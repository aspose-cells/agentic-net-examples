// Title: C# – Delete rows 60‑65 in an Excel file and convert to PDF with Aspose.Cells
// Description: Loads input.xlsx using Aspose.Cells, removes rows 60‑65 (zero‑based index 59, count 6) from the first worksheet, and saves the modified workbook directly as output.pdf via the Pdf SaveFormat.
// Keywords: Aspose.Cells DeleteRows | C# Excel delete rows | Aspose.Cells PDF conversion | remove rows 60-65 | SaveFormat.Pdf | Excel to PDF C#
// Common Searches: Aspose.Cells delete rows 60 to 65 C# | Convert Excel to PDF after removing rows with Aspose.Cells | How to delete a range of rows in Excel using Aspose.Cells .NET | C# Aspose.Cells remove rows and save as PDF
// Developer Intent: Remove rows 60‑65 from an Excel worksheet and export the updated workbook as a PDF file.
// Use Cases: Generate a clean PDF report by stripping out temporary or summary rows before conversion. | Exclude confidential or irrelevant rows from a financial sheet prior to archiving as PDF. | Prepare a printable version of a dataset after deleting blank or placeholder rows.
// AI Prompts: Write C# code with Aspose.Cells that deletes rows 60‑65 from the first worksheet and saves the workbook as a PDF. | Explain how to calculate the zero‑based index for a given Excel row number when using DeleteRows in Aspose.Cells. | Show error‑handling patterns for missing input files or out‑of‑range row deletions in Aspose.Cells PDF export.

using Aspose.Cells;
using Aspose.Cells.Utility;

// Loads input.xlsx using Aspose.Cells, removes rows 60‑65 (zero‑based index 59, count 6) from the first worksheet, and saves the modified workbook directly as output.pdf via the Pdf SaveFormat.
class Program
{
    static void Main()
    {
        // Load the existing Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Delete rows 60 through 65 (zero‑based index 59, total 6 rows)
        workbook.Worksheets[0].Cells.DeleteRows(59, 6);

        // Save the result directly as a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
