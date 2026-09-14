// Title: Conditionally freeze the header row in Excel worksheets that exceed 100 rows using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that loops through every worksheet and invokes sheet.FreezePanes(1,0,1,0) only when the sheet’s MaxDataRow is 100 or greater. | Update an existing Aspose.Cells program to add a row‑count check before applying FreezePanes, ensuring the header stays visible only on large worksheets.
// Common Searches: Aspose.Cells C# freeze top row when worksheet has more than 100 rows | Conditional FreezePanes based on row count in .NET Excel library | How to keep header visible for large Excel sheets using Aspose.Cells | C# example to apply freeze panes only to sheets with over a hundred rows
// Tags: conditional pane freeze Aspose.Cells .NET | freeze header Excel Aspose.Cells | MaxDataRow row count check Aspose.Cells | freeze panes large worksheets C# | Excel sheet row threshold Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program loads an existing workbook or creates a new one, iterates through each worksheet, checks the maximum used row index, and if a sheet contains 101 or more rows it freezes the first row with FreezePanes(1,0,1,0). The modified workbook is then saved as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Ensure the input file exists; if not, create a new workbook.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
            }

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine the last used row index (0‑based). -1 means the sheet is empty.
                int lastUsedRow = sheet.Cells.MaxDataRow;

                // If the worksheet contains more than 100 rows (i.e., index >= 100)
                if (lastUsedRow >= 100)
                {
                    // Apply freeze panes. Freeze the first row (row index 0) so the header stays visible.
                    // Parameters: row index to start scrolling, column index to start scrolling,
                    // number of rows to freeze, number of columns to freeze.
                    sheet.FreezePanes(1, 0, 1, 0);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
