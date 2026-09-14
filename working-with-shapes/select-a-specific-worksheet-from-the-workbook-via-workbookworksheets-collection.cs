// Title: How to select a worksheet by name or index in an Aspose.Cells workbook using C#
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, retrieves a worksheet by its name, and saves the workbook. | Show an example of accessing a worksheet by zero‑based index in Aspose.Cells for .NET and then performing a simple operation. | Provide a snippet that selects a specific sheet, modifies a cell, and saves the changes using Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# get worksheet by sheet name | select worksheet by index Aspose.Cells .NET example | how to access a specific sheet in a workbook using Aspose.Cells for C# | C# Aspose.Cells retrieve sheet and save workbook | load Excel file and choose sheet number with Aspose.Cells
// Tags: retrieve worksheet by name Aspose.Cells | access worksheet by index C# | load and save Excel workbook Aspose.Cells | select specific sheet in .NET workbook | Aspose.Cells worksheet selection example

using Aspose.Cells;

// // Loads input.xlsx, selects the worksheet named "Sheet2" (or by zero‑based index), optionally manipulates it, and saves the workbook to output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load an existing workbook (using the provided load rule)
        Workbook workbook = new Workbook("input.xlsx");

        // Select a worksheet by its name
        Worksheet selectedSheet = workbook.Worksheets["Sheet2"];

        // Alternatively, select by zero‑based index
        // Worksheet selectedSheet = workbook.Worksheets[1];

        // (Optional) Perform operations on the selected worksheet here

        // Save the workbook (using the provided save rule)
        workbook.Save("output.xlsx");
    }
}
