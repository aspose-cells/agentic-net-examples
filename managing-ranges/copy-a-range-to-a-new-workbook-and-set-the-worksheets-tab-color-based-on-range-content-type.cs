// Title: Copy a Range to a New Workbook and Color the Worksheet Tab Based on Numeric Content (C# Aspose.Cells)
// Description: Creates a source workbook, fills cells A1:B3, copies that range to D5:E7 in a new workbook, checks the source range for numeric values, sets the destination worksheet tab to green if numbers are found (otherwise red), and saves the file as an XLSX document using Aspose.Cells for .NET.
// Keywords: Aspose.Cells copy range | C# copy range to another workbook | worksheet tab color Aspose.Cells | detect numeric cells Aspose.Cells | range content type detection | preserve formatting when copying ranges | Excel automation C# | set tab color programmatically | Aspose.Cells Range.Copy example
// Common Searches: How to copy a cell range to a different workbook with Aspose.Cells C# | Set worksheet tab color based on cell values using Aspose.Cells | Detect numeric data in a range with Aspose.Cells .NET | Copy range with formatting and formulas Aspose.Cells | Change Excel sheet tab color programmatically C#
// Developer Intent: Copy a defined cell range into a new workbook and automatically change the destination worksheet's tab color according to whether the range contains numeric data.
// Use Cases: Create a summary report where sheets containing numeric metrics are highlighted with a green tab for quick identification. | Automate generation of separate workbooks for data blocks, using red or green tabs to signal textual versus numeric content. | Copy a data block to a template workbook while preserving formulas and styles, then flag the sheet with a colored tab based on the presence of numbers.
// AI Prompts: Generate C# code with Aspose.Cells that copies a range from one workbook to another and sets the destination worksheet tab to green if any cell is numeric, otherwise red. | Show how to iterate through a copied range in Aspose.Cells to detect numeric values and apply conditional tab colors. | Explain how to preserve formulas, formatting, and cell styles when using Aspose.Cells Range.Copy, then change the worksheet tab color based on the range's data type.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsRangeCopyDemo
{
    // Creates a source workbook, fills cells A1:B3, copies that range to D5:E7 in a new workbook, checks the source range for numeric values, sets the destination worksheet tab to green if numbers are found (otherwise red), and saves the file as an XLSX document using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create source workbook and populate a range ----------
                Workbook sourceWb = new Workbook();                     // create source workbook
                Worksheet sourceWs = sourceWb.Worksheets[0];            // get first worksheet

                // Fill some sample data in the source range A1:B3
                sourceWs.Cells["A1"].PutValue("Item");
                sourceWs.Cells["B1"].PutValue("Quantity");
                sourceWs.Cells["A2"].PutValue("Apple");
                sourceWs.Cells["B2"].PutValue(10);
                sourceWs.Cells["A3"].PutValue("Banana");
                sourceWs.Cells["B3"].PutValue(20);

                // Define the source range to be copied
                Aspose.Cells.Range sourceRange = sourceWs.Cells.CreateRange("A1:B3");

                // ---------- Create destination workbook ----------
                Workbook destWb = new Workbook();                       // create destination workbook
                Worksheet destWs = destWb.Worksheets[0];                // get first worksheet

                // Define the destination range where the source range will be copied
                // Here we start at cell D5, but any address works
                Aspose.Cells.Range destRange = destWs.Cells.CreateRange("D5:E7");

                // ---------- Copy the range ----------
                // The Copy method copies data, formulas, formatting, etc. from sourceRange to destRange
                destRange.Copy(sourceRange);

                // ---------- Determine content type of the source range ----------
                // Simple logic: if any cell in the range contains a numeric value, treat as "numeric"
                bool hasNumeric = false;
                foreach (Cell cell in sourceRange)
                {
                    if (cell.Type == CellValueType.IsNumeric)
                    {
                        hasNumeric = true;
                        break;
                    }
                }

                // ---------- Set worksheet tab color based on content type ----------
                // Numeric content -> Green tab, otherwise -> Red tab
                destWs.TabColor = hasNumeric ? Color.Green : Color.Red;

                // ---------- Save the destination workbook ----------
                destWb.Save("RangeCopyWithTabColor.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
