// Title: How to set print title rows that include frozen rows and confirm they stay visible when printing using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a workbook, freezes the first two rows, assigns those rows to the PrintTitleRows property, and saves the file with Aspose.Cells. | Write a C# snippet that reads the PrintTitleRows value after applying FreezePanes and writes it to the console to verify the setting.
// Common Searches: Aspose.Cells C# set PrintTitleRows for frozen rows | How to keep header rows visible when printing after FreezePanes in Aspose.Cells | C# example to verify PrintTitleRows property after applying FreezePanes | Print title rows with frozen panes using Aspose.Cells for .NET | Saving Excel file with print titles that include frozen rows in C#
// Tags: Aspose.Cells set PrintTitleRows property | Aspose.Cells freeze panes with print titles | C# print title rows after FreezePanes | Excel workbook print titles frozen rows Aspose.Cells | Aspose.Cells verify PrintTitleRows value

using Aspose.Cells;
using System;

// The example creates a new workbook, fills it with sample data, freezes the first two rows, sets the PrintTitleRows property to "$1:$2" so the frozen rows act as print titles, outputs the configured range to the console for verification, and saves the workbook as PrintTitleWithFrozenRows.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DataSheet";

            // Fill the worksheet with sample data
            for (int row = 0; row < 30; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Freeze the first two rows (rows 1 and 2 in Excel UI)
            // Parameters: row index after frozen rows, column index after frozen columns,
            // number of rows to freeze, number of columns to freeze
            sheet.FreezePanes(2, 0, 2, 0);

            // Apply print title range that includes the frozen rows
            // "$1:$2" means rows 1 through 2
            sheet.PageSetup.PrintTitleRows = "$1:$2";

            // Verify that the title range is set correctly
            string titleRange = sheet.PageSetup.PrintTitleRows;
            Console.WriteLine($"Print title rows set to: {titleRange}");

            // Save the workbook
            workbook.Save("PrintTitleWithFrozenRows.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
