// Title: How to set a custom column width for a specific table column in Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to set column B width to 30 characters. | Show how to apply the same column width to a consecutive range of columns (e.g., columns C‑G) with Aspose.Cells. | Demonstrate saving a workbook after adjusting column widths using Aspose.Cells in C#. | Explain how to choose an appropriate character width for long text strings with the SetColumnWidth method.
// Common Searches: Aspose.Cells C# set column width to fit long text | how to change column width for a specific column in a worksheet using Aspose.Cells | set same column width for multiple columns Aspose.Cells example | Aspose.Cells SetColumnWidth characters parameter usage | adjust column width for table column in Aspose.Cells .NET
// Tags: Aspose.Cells SetColumnWidth method | custom column width Aspose.Cells C# | adjust worksheet column width Aspose.Cells | uniform column width range Aspose.Cells | column width characters Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a workbook, accesses the first worksheet, sets column B's width to 30 characters with SetColumnWidth, shows how to apply the same width to a range of columns, and saves the file as CustomColumnWidth.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (you can also load an existing file with new Workbook("input.xlsx"))
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set a custom width for column B (zero‑based index 1)
        // Width is expressed in characters of the default font
        sheet.Cells.SetColumnWidth(1, 30); // 30 characters wide to fit long text

        // If you need to set the same width for a range of columns, use the overload:
        // sheet.Cells.SetColumnWidth(startColumn, totalColumns, width);
        // Example: sheet.Cells.SetColumnWidth(2, 5, 20); // columns C‑G each 20 characters wide

        // Save the workbook to a file
        workbook.Save("CustomColumnWidth.xlsx");
    }
}
