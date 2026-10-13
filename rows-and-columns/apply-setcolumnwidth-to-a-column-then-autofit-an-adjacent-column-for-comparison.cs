// Title: Set a fixed width for column A and auto‑fit column B using Aspose.Cells for .NET (C#)
// AI Prompts: Create a C# program that builds a new workbook, assigns a width of 20 characters to column A, automatically adjusts column B to its content, and writes the result to ColumnWidthComparison.xlsx using Aspose.Cells. | Write C# code with Aspose.Cells that demonstrates applying a fixed column width to one column and then invoking the auto‑fit routine on the neighboring column before saving.
// Common Searches: asp.net aspose.cells set column width to specific number of characters | how to auto adjust column width after setting custom width in Aspose.Cells C# | c# example of comparing fixed column width with auto‑fit in Aspose.Cells | Aspose.Cells column width vs auto‑fit behavior demonstration
// Tags: specific column width Aspose.Cells C# | auto‑fit column Aspose.Cells C# | column width comparison Aspose.Cells | C# worksheet column sizing Aspose.Cells | Aspose.Cells column width manipulation example

using Aspose.Cells;

// Demonstrates creating a workbook, setting column A to a 20‑character width, auto‑fitting column B based on its content, and saving the file as ColumnWidthComparison.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set a specific width (in characters) for column A (index 0)
        sheet.Cells.SetColumnWidth(0, 20); // 20 characters wide

        // Auto‑fit column B (index 1) based on its content
        sheet.AutoFitColumn(1);

        // Save the workbook to a file
        workbook.Save("ColumnWidthComparison.xlsx");
    }
}
