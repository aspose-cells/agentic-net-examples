// Title: How to prepend an apostrophe to numeric cells in Excel using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add leading apostrophe to numeric values in a worksheet | convert numeric cells to text with an apostrophe using Aspose.Cells .NET | force Excel to treat numbers as text programmatically with Aspose.Cells | iterate over cells and prefix numbers with an apostrophe in Aspose.Cells C# | save workbook after changing cell type to text with Aspose.Cells
// Tags: Cell.PutValue with leading apostrophe Aspose.Cells | CellValueType.IsNumeric check Aspose.Cells | Iterate over Cells collection to modify values Aspose.Cells | Convert numeric cell to text in Excel using Aspose.Cells | Workbook.Save after text conversion Aspose.Cells .NET

using System;
using Aspose.Cells;

namespace AsposeCellsApostrophePrefix
{
    // The example loads an existing workbook, iterates through every used cell, detects numeric values via CellValueType.IsNumeric, prefixes each with an apostrophe using Cell.PutValue to force text interpretation, and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("Input.xlsx");

            // Get the first worksheet (or iterate through all worksheets as needed)
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Iterate through all used cells in the worksheet
            foreach (Cell cell in cells)
            {
                // Check if the cell contains a numeric value (int, double, decimal, etc.)
                if (cell.Type == CellValueType.IsNumeric)
                {
                    // Retrieve the original numeric value as string
                    string originalValue = cell.Value.ToString();

                    // Prefix with an apostrophe to force text interpretation
                    // The apostrophe itself is not displayed in Excel, but forces the cell to be treated as text
                    cell.PutValue("'" + originalValue);
                }
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("Output.xlsx");
        }
    }
}
