// Title: How to apply a custom number format with thousand separators and no decimal places to Excel cells using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set cell format to display numbers with commas and no decimals | C# Aspose.Cells custom number format '#,##0' for an Excel column | Formatting numeric cells with thousand separators in Aspose.Cells .NET | Apply number style without fractional part using Aspose.Cells API | Save workbook after applying custom number format in Aspose.Cells C#
// Tags: custom number format Aspose.Cells C# | thousand separator cell style Aspose.Cells | apply numeric format without decimals Aspose.Cells | Excel number formatting Aspose.Cells .NET | set cell style custom format Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a workbook, writes quantity values to a worksheet, applies the custom number format "#,##0" to each cell to show thousand separators with no decimal places, and saves the file as QuantitiesFormatted.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one if needed)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Quantities";

        // Sample data: quantities to display
        double[] quantities = { 1234, 56789, 1234567, 89012345 };

        // Write the quantities to column A starting from row 1
        for (int i = 0; i < quantities.Length; i++)
        {
            // Write raw numeric value
            Cell cell = sheet.Cells[i, 0];
            cell.PutValue(quantities[i]);

            // Apply custom number format: thousand separators, no decimals
            Style style = cell.GetStyle();
            style.Custom = "#,##0";   // e.g., 1,234 ; 56,789 ; 1,234,567
            cell.SetStyle(style);
        }

        // Save the workbook to a file
        workbook.Save("QuantitiesFormatted.xlsx");
    }
}
