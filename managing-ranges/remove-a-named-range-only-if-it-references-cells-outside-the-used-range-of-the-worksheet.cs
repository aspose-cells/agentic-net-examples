// Title: Remove named ranges outside the used range with Aspose.Cells for .NET
// Description: C# example that identifies named ranges whose referenced cells exceed the worksheet's MaxDataRow/MaxDataColumn and deletes them using Aspose.Cells.
// Keywords: Aspose.Cells remove named range | delete out‑of‑bounds named range .NET | GetRange MaxDataRow MaxDataColumn | Workbook.Worksheets.Names.Remove | Excel named range cleanup
// Common Searches: how to delete named ranges outside used area Aspose.Cells | remove invalid named ranges .NET Excel library | filter named ranges by used range Aspose.Cells | Aspose.Cells find and delete out of range names
// Developer Intent: Automatically purge any named range that points to cells beyond the worksheet's actual data area.
// Use Cases: Sanitize a workbook before sharing to eliminate references to empty cells. | Validate imported Excel files and strip out‑of‑bounds names to avoid runtime errors. | Prepare a data‑entry template, keeping only ranges that intersect the populated region.
// AI Prompts: Write C# code with Aspose.Cells that scans Workbook.Worksheets.Names, checks each name's GetRange() against ws.Cells.MaxDataRow and MaxDataColumn, and removes the out‑of‑range entries. | Provide a snippet that logs the names of all ranges removed because they fall outside the used range before saving the file. | Explain how to extend the logic to skip names that refer to formulas, external workbooks, or whole‑column/whole‑row references.

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// C# example that identifies named ranges whose referenced cells exceed the worksheet's MaxDataRow/MaxDataColumn and deletes them using Aspose.Cells.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (replace with new Workbook("input.xlsx") to load an existing file)
            Workbook workbook = new Workbook();

            // Example data to define a used range
            Worksheet ws = workbook.Worksheets[0];
            ws.Cells["A1"].PutValue(1);
            ws.Cells["B2"].PutValue(2);

            // Add a named range that lies inside the used range
            int insideIdx = workbook.Worksheets.Names.Add("InsideRange");
            workbook.Worksheets.Names[insideIdx].RefersTo = "=Sheet1!$A$1:$B$2";

            // Add a named range that lies outside the used range
            int outsideIdx = workbook.Worksheets.Names.Add("OutsideRange");
            workbook.Worksheets.Names[outsideIdx].RefersTo = "=Sheet1!$Z$100:$AA$101";

            // Determine the used range limits (zero‑based indices)
            int maxRow = ws.Cells.MaxDataRow;
            int maxCol = ws.Cells.MaxDataColumn;

            // Collect names that reference cells outside the used range
            List<string> namesToRemove = new List<string>();
            foreach (Name name in workbook.Worksheets.Names)
            {
                // Get the range the name refers to
                Aspose.Cells.Range rng = name.GetRange();
                if (rng == null) continue; // Skip if the name does not refer to a range

                int firstRow = rng.FirstRow;
                int firstCol = rng.FirstColumn;
                int lastRow = firstRow + rng.RowCount - 1;
                int lastCol = firstCol + rng.ColumnCount - 1;

                // If any part of the range is outside the used range, mark it for removal
                if (firstRow > maxRow || firstCol > maxCol || lastRow > maxRow || lastCol > maxCol)
                {
                    namesToRemove.Add(name.Text);
                }
            }

            // Remove the identified named ranges
            foreach (string nameText in namesToRemove)
            {
                workbook.Worksheets.Names.Remove(nameText);
            }

            // Save the workbook
            string outputPath = "Result.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
