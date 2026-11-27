// Title: Split worksheet panes at row 10 and column 5 with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# split worksheet panes at specific row and column | set split rows and split columns programmatically using Aspose.Cells | how to create frozen panes at row 10 column 5 in an Excel file with Aspose.Cells | using reflection to access View properties in Aspose.Cells C# example | save Excel workbook after configuring split panes with Aspose.Cells .NET
// Tags: aspacells split worksheet panes | aspacells set split rows property | aspacells set split columns property | c# excel split panes aspacells | aspacells worksheet view configuration

using System;
using System.IO;
using Aspose.Cells;

namespace SplitPaneExampleApp
{
    // The example creates a new Workbook, accesses the first Worksheet, and uses reflection to reach the hidden View object. It sets SplitRows to 9 (row 10) and SplitColumns to 4 (column 5), optionally defines the ActiveCell, then saves the file as SplitPaneResult.xlsx.
    class SplitPaneExample
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Attempt to configure split panes via reflection (covers versions without direct View property)
                var viewProp = sheet.GetType().GetProperty("View");
                if (viewProp != null)
                {
                    object viewObj = viewProp.GetValue(sheet);
                    if (viewObj != null)
                    {
                        var splitRowsProp = viewObj.GetType().GetProperty("SplitRows");
                        var splitColumnsProp = viewObj.GetType().GetProperty("SplitColumns");
                        var activeCellProp = viewObj.GetType().GetProperty("ActiveCell");

                        splitRowsProp?.SetValue(viewObj, 9);      // rows above the split line
                        splitColumnsProp?.SetValue(viewObj, 4);   // columns to the left of the split line
                        activeCellProp?.SetValue(viewObj, "A1"); // optional active cell
                    }
                }

                // Define output file path
                string outputPath = "SplitPaneResult.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
