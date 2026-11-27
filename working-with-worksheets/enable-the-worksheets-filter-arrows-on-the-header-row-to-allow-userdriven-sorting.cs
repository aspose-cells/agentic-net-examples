// Title: How to enable AutoFilter arrows on an Excel worksheet header row using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set AutoFilter range for header row to show filter arrows | Enable Excel column filter dropdowns programmatically with Aspose.Cells .NET | Add AutoFilter to the first row of a workbook using Aspose.Cells in C# | Create a sortable Excel sheet with filter arrows via Aspose.Cells API | Determine last used column and set worksheet.AutoFilter.Range dynamically in Aspose.Cells
// Tags: Aspose.Cells AutoFilter header range C# | enable filter arrows worksheet Aspose.Cells | dynamic AutoFilter range calculation Aspose.Cells | programmatic Excel sorting arrows Aspose.Cells | C# calculate last column for AutoFilter Aspose

using Aspose.Cells;
using System;

// The example creates a new workbook, adds sample data with a header row, computes the last used column to build the header range string, assigns that range to the worksheet's AutoFilter, refreshes it to display filter arrows for user‑driven sorting, and saves the file as FilteredWorkbook.xlsx.
class EnableFilterArrows
{
    static void Main()
    {
        // Create a new workbook (lifecycle create rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data with a header row
        sheet.Cells["A1"].PutValue("ID");
        sheet.Cells["B1"].PutValue("Name");
        sheet.Cells["C1"].PutValue("Category");
        sheet.Cells["D1"].PutValue("Price");

        sheet.Cells["A2"].PutValue(1);
        sheet.Cells["B2"].PutValue("Apple");
        sheet.Cells["C2"].PutValue("Fruit");
        sheet.Cells["D2"].PutValue(0.5);

        sheet.Cells["A3"].PutValue(2);
        sheet.Cells["B3"].PutValue("Carrot");
        sheet.Cells["C3"].PutValue("Vegetable");
        sheet.Cells["D3"].PutValue(0.3);

        // Determine the last used column in the header row
        int lastColumn = sheet.Cells.MaxColumn;

        // Build the range string for the header row (e.g., "A1:D1")
        string headerRange = $"A1:{CellIndexToName(0, lastColumn)}1";

        // Enable filter arrows on the header row (feature implementation)
        sheet.AutoFilter.Range = headerRange;
        sheet.AutoFilter.Refresh();

        // Save the workbook (lifecycle save rule)
        workbook.Save("FilteredWorkbook.xlsx");
    }

    // Helper method to convert column index to Excel column name (0‑based index)
    private static string CellIndexToName(int rowIndex, int columnIndex)
    {
        int dividend = columnIndex + 1;
        string columnName = string.Empty;
        while (dividend > 0)
        {
            int modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar(65 + modulo) + columnName;
            dividend = (dividend - modulo) / 26;
        }
        // Row index is not used for column name conversion; return column name only
        return columnName;
    }
}
