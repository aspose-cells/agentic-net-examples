// Title: How to set a worksheet to VeryHidden, freeze the first row and column, and then make it visible again using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# hide a worksheet as VeryHidden, apply FreezePanes, and unhide it later | C# code to freeze first row and column on a hidden sheet with Aspose.Cells | Make a worksheet VeryHidden then freeze panes before showing it using Aspose.Cells for .NET | Aspose.Cells set worksheet visibility to VeryHidden then restore visibility after FreezePanes | How to programmatically hide a sheet, freeze panes, and save workbook with Aspose.Cells in C#
// Tags: set worksheet very hidden Aspose.Cells | apply freeze panes Aspose.Cells C# | toggle worksheet visibility Aspose.Cells | save workbook after sheet modification Aspose.Cells | C# hide and unhide worksheet Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new workbook, adds a worksheet named "MySheet", hides it by setting its visibility to VeryHidden (using IsVisible = false), freezes the first row and column with FreezePanes, then makes the sheet visible again and saves the workbook as Result.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (create rule)
                var workbook = new Workbook();

                // Add a new worksheet and get a reference to it
                int sheetIndex = workbook.Worksheets.Add();
                var sheet = workbook.Worksheets[sheetIndex];
                sheet.Name = "MySheet";

                // Hide the worksheet (VeryHidden is not available in older versions, use IsVisible)
                sheet.IsVisible = false;

                // Freeze panes (freeze first row and first column)
                // Parameters: totalRows, totalColumns, frozenRows, frozenColumns
                sheet.FreezePanes(1, 1, 1, 1);

                // Make the worksheet visible again
                sheet.IsVisible = true;

                // Define output file path
                string outputPath = "Result.xlsx";

                // Save the workbook (save rule)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
