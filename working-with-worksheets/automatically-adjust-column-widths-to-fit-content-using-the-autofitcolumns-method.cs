// Title: Automatically adjust Excel column widths to content with Aspose.Cells AutoFitColumns in C#
// AI Prompts: Generate a C# example that opens an existing .xlsx file with Aspose.Cells, calls AutoFitColumns on the first worksheet, and saves the result. | Write a C# method that receives a Worksheet object and auto‑fits a specified range of columns (e.g., columns B to E) using Aspose.Cells. | Create a C# snippet that adds new rows to a worksheet, then invokes AutoFitColumns to re‑size all columns before saving. | Provide C# code that loads a workbook, disables gridlines, auto‑fits columns, and exports the sheet as a PDF using Aspose.Cells.
// Common Searches: Aspose.Cells C# auto fit column width after loading workbook | How to use AutoFitColumns for a specific column range in Aspose.Cells .NET | Resize all Excel columns to fit content with Aspose.Cells in C# | C# code example for AutoFitColumns on a worksheet using Aspose.Cells | AutoFitColumns method performance considerations in Aspose.Cells .NET
// Tags: auto-fit columns Aspose.Cells .NET | AutoFitColumns worksheet C# | adjust Excel column width Aspose.Cells | auto-fit column range Aspose.Cells | resize columns after adding rows Aspose.Cells

using Aspose.Cells;

// // Loads 'input.xlsx', automatically adjusts the width of every column in the first worksheet with AutoFitColumns, and saves the modified workbook as 'output.xlsx'.
class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Auto-fit all columns to match the content
        sheet.AutoFitColumns();

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
