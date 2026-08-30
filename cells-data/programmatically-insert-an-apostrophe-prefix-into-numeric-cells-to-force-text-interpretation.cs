// Title: Add an apostrophe prefix to numeric cells to store them as text using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loops through a worksheet and enables the QuotePrefix flag on each numeric cell to force text storage. | Create a reusable style with QuotePrefix set to true and apply it via StyleFlag to numeric cells in an Aspose.Cells workbook. | Write a C# program that saves an Excel file where numbers appear with a leading single quote by using Aspose.Cells SetStyle method.
// Common Searches: Aspose.Cells C# how to add a leading single quote to numeric cells | set QuotePrefix on cells programmatically with Aspose.Cells .NET | convert numbers to text in Excel using Aspose.Cells style flag | force Excel numeric values to be treated as text in C# Aspose.Cells
// Tags: Aspose.Cells QuotePrefix style flag C# | C# Aspose.Cells numeric cell as text | Apply StyleFlag QuotePrefix Aspose.Cells | Excel apostrophe prefix using Aspose.Cells .NET | SetStyle for numeric cells Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsQuotePrefixDemo
{
    // The example creates a workbook, writes numeric values to cells, iterates over all cells, and for each numeric cell creates a style with QuotePrefix enabled. It then applies the style using a StyleFlag and saves the file, resulting in numbers stored as text with a leading apostrophe.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some numeric values
            cells["A1"].PutValue(12345);
            cells["A2"].PutValue(67890);
            cells["B1"].PutValue(3.14159);
            cells["B2"].PutValue(2.71828);

            // Apply QuotePrefix style to each numeric cell to force text interpretation
            // The style flag enables the QuotePrefix property for the style being applied
            foreach (Cell cell in cells)
            {
                // Check if the cell currently holds a numeric value
                if (cell.Type == CellValueType.IsNumeric)
                {
                    // Create a new style based on the workbook's style collection
                    Style style = workbook.CreateStyle();
                    style.QuotePrefix = true; // Mark that the cell should start with a single quote

                    // Create a style flag and enable the QuotePrefix flag
                    StyleFlag flag = new StyleFlag();
                    flag.QuotePrefix = true;

                    // Apply the style with the flag (SetStyle, not ApplyStyle)
                    cell.SetStyle(style, flag);
                }
            }

            // Save the workbook (lifecycle save)
            workbook.Save("NumericCellsWithApostrophePrefix.xlsx");
        }
    }
}
