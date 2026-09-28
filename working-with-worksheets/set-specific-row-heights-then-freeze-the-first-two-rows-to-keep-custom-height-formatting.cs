// Title: How to set custom row heights and freeze the first two rows in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that assigns a height of 30 points to row 1, 45 points to row 2, freezes the top two rows, and saves the workbook as an .xlsx file. | Create a .NET example that applies specific row heights, then calls FreezePanes to lock the first two rows while keeping all columns scrollable.
// Common Searches: Aspose.Cells C# set row height 30 points and freeze top two rows | C# example for freezing first two rows after adjusting row heights with Aspose.Cells | How to use FreezePanes to lock header rows in an Excel file via Aspose.Cells .NET | Set custom row heights and freeze panes in Excel using Aspose.Cells for .NET
// Tags: row height adjustment Aspose.Cells C# | freeze top rows Aspose.Cells .NET | custom row heights Excel Aspose.Cells | freeze panes after row formatting Aspose.Cells | Aspose.Cells worksheet row formatting example

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // // Creates a new workbook, sets the first row height to 30 points and the second row height to 45 points, freezes those two rows with FreezePanes, and saves the file as CustomRowHeightAndFreeze.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Set custom heights for the first two rows (indices 0 and 1)
                sheet.Cells.SetRowHeight(0, 30); // Row 1 height = 30 points
                sheet.Cells.SetRowHeight(1, 45); // Row 2 height = 45 points

                // Freeze the first two rows (rows 0 and 1) while keeping all columns scrollable
                // FreezePanes(row, column, totalRows, totalColumns)
                sheet.FreezePanes(2, 0, 0, 0);

                // Save the workbook
                string outputPath = "CustomRowHeightAndFreeze.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
