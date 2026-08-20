// Title: C# – Preserve Header Row Formatting When Converting an Aspose.Cells Table to a Range (TableToRangeOptions)
// Description: Demonstrates how to create a workbook, style a table header, and convert the ListObject to a plain range while retaining only the header row’s formatting by using TableToRangeOptions.LastRow = 0. The result is saved as an Excel file.
// Keywords: Aspose.Cells | TableToRangeOptions | preserve header formatting | ConvertToRange | C# | ListObject to range | Excel table conversion | header style retention | Aspose.Cells .NET example | range conversion formatting
// Common Searches: Aspose.Cells keep header style when converting table to range | TableToRangeOptions LastRow usage C# | convert ListObject to range without losing formatting | preserve table header formatting Aspose.Cells | C# example TableToRangeOptions header only
// Developer Intent: Convert a ListObject to a regular range while retaining only the header row’s formatting.
// Use Cases: Generate a report where calculations use a table, then export a clean range with a styled header. | Apply custom styling to a table header, flatten the table for distribution, and keep the header appearance. | Create a data‑entry template with a table, then produce a final workbook that contains only the formatted header as a normal range.
// AI Prompts: Show how to use TableToRangeOptions in Aspose.Cells for C# to keep only the first row’s formatting when converting a ListObject to a range. | Provide a C# code snippet that styles a table header and then calls ConvertToRange with LastRow = 0 to preserve that style. | Explain the impact of setting TableToRangeOptions.LastRow = 0 on formatting during a table‑to‑range conversion in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, style a table header, and convert the ListObject to a plain range while retaining only the header row’s formatting by using TableToRangeOptions.LastRow = 0. The result is saved as an Excel file.
    public class PreserveHeaderFormattingDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate header row
                cells["A1"].PutValue("ID");
                cells["B1"].PutValue("Name");
                cells["C1"].PutValue("Score");

                // Populate some data rows
                for (int row = 2; row <= 5; row++)
                {
                    cells[row - 1, 0].PutValue(row - 1);                     // ID
                    cells[row - 1, 1].PutValue($"Person {row - 1}");        // Name
                    cells[row - 1, 2].PutValue((row - 1) * 10);             // Score
                }

                // Add a table that includes the header and data rows (A1:C5)
                int tableIndex = sheet.ListObjects.Add(0, 0, 4, 2, true);
                ListObject table = sheet.ListObjects[tableIndex];

                // Apply a distinct style to the header row of the table
                Style headerStyle = workbook.CreateStyle();
                headerStyle.Font.IsBold = true;
                headerStyle.ForegroundColor = Color.LightBlue;
                headerStyle.Pattern = BackgroundType.Solid;

                // Optional built‑in table style (removed TableStyle property which is unavailable in some versions)
                table.ShowHeaderRow = true;
                table.ShowTableStyleFirstColumn = false;
                table.ShowTableStyleLastColumn = false;

                // Directly set the style for the header row range
                Aspose.Cells.Range headerRange = sheet.Cells.CreateRange("A1:C1");
                headerRange.SetStyle(headerStyle);

                // Convert the table to a range while preserving only the header row formatting.
                // Setting LastRow = 0 tells the conversion to keep rows up to index 0 (the header).
                TableToRangeOptions options = new TableToRangeOptions
                {
                    LastRow = 0   // preserve header row only
                };
                table.ConvertToRange(options);

                // Save the workbook
                workbook.Save("PreserveHeaderFormatting.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main()
        {
            PreserveHeaderFormattingDemo.Run();
        }
    }
}
