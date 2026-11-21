// Title: How to display a table header row and lock the header cells in an Excel file using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# show table header row and protect header cells | set ShowHeaderRow true and lock header row in Excel with Aspose.Cells | C# lock ListObject header cells and protect worksheet using Aspose.Cells | enable table header and apply locked style with Aspose.Cells .NET | protect Excel worksheet after locking table header in Aspose.Cells
// Tags: Aspose.Cells ShowHeaderRow property | apply locked style to Excel table header | protect worksheet after styling cells Aspose.Cells | C# lock header row of ListObject | Excel table header protection with Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The example loads an existing workbook, ensures a table (ListObject) is present, sets the table's ShowHeaderRow property to true, creates a locked style, applies it to each header cell, protects the worksheet with full protection, and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one table (ListObject)
            if (worksheet.ListObjects.Count == 0)
            {
                Console.WriteLine("No tables (ListObjects) found in the worksheet.");
                return;
            }

            // Get the first table (ListObject) on the worksheet
            ListObject table = worksheet.ListObjects[0];

            // Ensure the header row is displayed
            table.ShowHeaderRow = true;

            // Determine header row and column indices
            int headerRowIndex = table.StartRow;                     // Header row index
            int firstColumnIndex = table.StartColumn;               // First column index
            int lastColumnIndex = table.EndColumn;                  // Last column index

            // Create a style that locks cells
            Style lockedStyle = workbook.CreateStyle();
            lockedStyle.IsLocked = true;

            // Define which style attributes to apply (only the lock flag)
            StyleFlag styleFlag = new StyleFlag
            {
                Locked = true
            };

            // Apply the locked style to each cell in the header row
            for (int col = firstColumnIndex; col <= lastColumnIndex; col++)
            {
                Cell headerCell = worksheet.Cells[headerRowIndex, col];
                headerCell.SetStyle(lockedStyle, styleFlag);
            }

            // Protect the worksheet so that locked cells cannot be edited
            worksheet.Protect(ProtectionType.All);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
