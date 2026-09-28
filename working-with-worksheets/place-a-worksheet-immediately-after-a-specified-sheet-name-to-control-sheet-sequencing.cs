// Title: Insert a newly created worksheet immediately after a given sheet name with Aspose.Cells for .NET
// AI Prompts: Create a worksheet called 'NewSheet' and place it directly after the existing worksheet 'Sheet1' in a workbook using Aspose.Cells. | Reorder worksheets by removing a newly added sheet and inserting it at the index that follows a target sheet identified by its name, then save the workbook.
// Common Searches: Aspose.Cells C# insert worksheet after existing sheet by name | how to move a newly added worksheet to follow a specific worksheet in Aspose.Cells | reposition worksheet to a particular index using Aspose.Cells .NET | programmatically set worksheet order based on sheet name Aspose.Cells
// Tags: insert worksheet after target sheet Aspose.Cells | reorder worksheets by name .NET | add worksheet at specific index C# Aspose.Cells | move worksheet to desired position Aspose.Cells | worksheet sequencing with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The code loads or creates a workbook, adds a worksheet named 'NewSheet', locates the worksheet named 'Sheet1', removes the new sheet and inserts it immediately after 'Sheet1', and finally saves the workbook as output.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                Workbook workbook;

                // Load existing workbook if present, otherwise create a new one
                if (File.Exists(inputPath))
                {
                    try
                    {
                        workbook = new Workbook(inputPath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to load '{inputPath}': {ex.Message}");
                        workbook = new Workbook();
                    }
                }
                else
                {
                    workbook = new Workbook();
                }

                // Add a new worksheet that we want to position
                int newSheetIndex = workbook.Worksheets.Add();
                Worksheet newSheet = workbook.Worksheets[newSheetIndex];
                newSheet.Name = "NewSheet";

                // Name of the sheet after which the new sheet should appear
                const string targetSheetName = "Sheet1";

                // Locate the target sheet by name (returns null if not found)
                Worksheet targetSheet = workbook.Worksheets[targetSheetName];

                if (targetSheet != null)
                {
                    int targetIndex = workbook.Worksheets.IndexOf(targetSheet);
                    int sourceIndex = workbook.Worksheets.IndexOf(newSheet);

                    // Reposition the new sheet to be immediately after the target sheet
                    try
                    {
                        // Remove the source sheet
                        workbook.Worksheets.RemoveAt(sourceIndex);

                        // Determine correct insertion index after removal
                        int insertIndex = sourceIndex < targetIndex ? targetIndex : targetIndex + 1;

                        // Insert a new worksheet at the desired position
                        Worksheet inserted = workbook.Worksheets.Insert(insertIndex, SheetType.Worksheet);
                        inserted.Name = newSheet.Name;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to reposition worksheet: {ex.Message}");
                    }
                }

                // Save the workbook
                const string outputPath = "output.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to '{outputPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
