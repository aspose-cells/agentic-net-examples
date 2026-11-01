// Title: Copy only visible rows from a source worksheet to a new sheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# copy rows while ignoring hidden rows | How to skip hidden rows when copying data between worksheets in .NET | Copy visible rows to another sheet using Aspose.Cells API | C# example for filtering hidden rows before copying with Aspose.Cells | Transfer only non‑hidden rows from one Excel sheet to another using Aspose.Cells | Aspose.Cells copy rows without hidden rows performance
// Tags: copy visible rows Aspose.Cells C# | exclude hidden rows during worksheet copy .NET | pre‑filter rows for CopyRows Aspose.Cells | transfer non‑hidden Excel rows C# | Aspose.Cells row visibility handling

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a source.xlsx (if missing) with a hidden third row, loads the workbook, adds a destination worksheet, iterates through each source row, skips rows where IsHidden is true, copies each cell's value and style to the new sheet, and saves the result as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the source file exists; create a simple one if missing.
            if (!File.Exists(sourcePath))
            {
                var tempWb = new Workbook();
                var tempSheet = tempWb.Worksheets[0];
                tempSheet.Name = "Sheet1";

                // Populate sample data
                tempSheet.Cells["A1"].PutValue("Header");
                tempSheet.Cells["A2"].PutValue("Row 1");
                tempSheet.Cells["A3"].PutValue("Row 2");
                // Hide row 2 as an example (row index 2 corresponds to third row)
                tempSheet.Cells.Rows[2].IsHidden = true;

                tempWb.Save(sourcePath);
            }

            // Load the source workbook
            var workbook = new Workbook(sourcePath);

            // Get the source worksheet by name
            var srcSheet = workbook.Worksheets["Sheet1"];
            if (srcSheet == null)
                throw new InvalidOperationException("Worksheet 'Sheet1' not found in the source workbook.");

            // Add a new worksheet to receive the copied rows
            var destSheet = workbook.Worksheets.Add("Copy");

            int destRowIndex = 0;
            int lastRow = srcSheet.Cells.MaxDataRow;
            int maxCol = srcSheet.Cells.MaxDataColumn;

            // Iterate through each row in the source sheet
            for (int srcRowIndex = 0; srcRowIndex <= lastRow; srcRowIndex++)
            {
                // Skip hidden rows
                if (srcSheet.Cells.Rows[srcRowIndex].IsHidden)
                    continue;

                // Copy each cell in the visible row to the destination sheet
                for (int col = 0; col <= maxCol; col++)
                {
                    var srcCell = srcSheet.Cells[srcRowIndex, col];
                    var destCell = destSheet.Cells[destRowIndex, col];

                    // Copy value
                    destCell.PutValue(srcCell.Value);

                    // Copy style (optional, can be omitted for performance)
                    destCell.SetStyle(srcCell.GetStyle());
                }

                destRowIndex++;
            }

            // Save the workbook with the copied rows
            workbook.Save(outputPath);
            Console.WriteLine($"Rows copied successfully. Output saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
