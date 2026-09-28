// Title: Hide columns C and D in an Excel workbook using Aspose.Cells for .NET and save while preserving hidden columns
// AI Prompts: Load an existing .xlsx file with Aspose.Cells, hide columns C and D, and save the workbook so the hidden state remains intact. | Programmatically conceal specific columns in a worksheet using Aspose.Cells for .NET and export the file with the hidden columns retained.
// Common Searches: asp.net hide column c and d using aspose.cells | preserve hidden columns when saving excel workbook with aspose.cells .net | aspose.cells hide multiple columns by index c# | how to keep columns hidden after exporting workbook with aspose.cells
// Tags: hide columns Aspose.Cells | preserve hidden column state on save | Aspose.Cells hide column by index | export workbook with hidden columns .NET

using System;
using Aspose.Cells;

// // Loads 'input.xlsx', hides columns C (index 2) and D (index 3) on the first worksheet using Aspose.Cells, then saves to 'output.xlsx' while retaining the hidden column settings.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (adjust index or name as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Hide column C (zero‑based index 2) and column D (index 3)
        sheet.Cells.HideColumn(2);
        sheet.Cells.HideColumn(3);

        // Save the workbook; hidden columns are preserved automatically
        workbook.Save("output.xlsx");
    }
}
