// Title: Remove duplicate entries from a named range with Aspose.Cells FindOptions (C#)
// Description: Demonstrates how to create a workbook, define a named range, and use Aspose.Cells FindOptions to locate and clear duplicate values while preserving the first occurrence. The example works with .NET (C#) and shows how to restrict searches to a specific named range, making the data list unique.
// Keywords: Aspose.Cells | C# | FindOptions | named range | remove duplicates | duplicate values | clear cells | Excel unique list | Aspose.Cells .NET | data cleaning
// Common Searches: Aspose.Cells remove duplicate values from named range | FindOptions limit search to specific range C# | How to clear duplicate cells in Excel using Aspose | Keep first occurrence of a value in Aspose.Cells | C# code to deduplicate a named range in Excel
// Developer Intent: The developer wants to keep only the first occurrence of each value inside a named range and erase all subsequent duplicates using Aspose.Cells FindOptions.
// Use Cases: Cleaning imported lists where duplicate entries must be eliminated but the original entry retained. | Preparing a named range for data‑validation rules that require unique items. | Generating reports that need a distinct set of values extracted from a predefined range. | Automating spreadsheet cleanup in multi‑regional deployments (US, EU, APAC) using a single .NET solution.
// AI Prompts: Show a compact Aspose.Cells .NET snippet that removes duplicate values from a named range using FindOptions. | Explain step‑by‑step how to configure FindOptions to restrict a Find operation to a specific named range and clear duplicate cells. | Suggest performance‑oriented improvements for the duplicate‑removal loop when processing large named ranges.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, define a named range, and use Aspose.Cells FindOptions to locate and clear duplicate values while preserving the first occurrence. The example works with .NET (C#) and shows how to restrict searches to a specific named range, making the data list unique.
    public class RemoveDuplicateEntriesInNamedRange
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data with duplicates in column A
                sheet.Cells["A1"].PutValue("Apple");
                sheet.Cells["A2"].PutValue("Banana");
                sheet.Cells["A3"].PutValue("Apple");   // duplicate
                sheet.Cells["A4"].PutValue("Cherry");
                sheet.Cells["A5"].PutValue("Banana");  // duplicate
                sheet.Cells["A6"].PutValue("Date");

                // Define a named range that covers the populated cells
                int nameIdx = workbook.Worksheets.Names.Add("MyRange");
                workbook.Worksheets.Names[nameIdx].RefersTo = "=Sheet1!$A$1:$A$6";

                // Retrieve the named range object
                Name namedRange = workbook.Worksheets.Names["MyRange"];
                Aspose.Cells.Range range = namedRange.GetRange();

                // Build a CellArea that represents the same range – required for FindOptions
                CellArea searchArea = new CellArea
                {
                    StartRow = range.FirstRow,
                    StartColumn = range.FirstColumn,
                    EndRow = range.FirstRow + range.RowCount - 1,
                    EndColumn = range.FirstColumn + range.ColumnCount - 1
                };

                // Configure FindOptions to limit searches to the named range
                FindOptions findOptions = new FindOptions();
                findOptions.SetRange(searchArea);
                findOptions.LookInType = LookInType.Values;
                // Default LookAtType is Whole, so no explicit setting is required

                // Keep track of values that have already been processed
                HashSet<string> processedValues = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                // Iterate through each cell in the named range
                for (int r = range.FirstRow; r < range.FirstRow + range.RowCount; r++)
                {
                    for (int c = range.FirstColumn; c < range.FirstColumn + range.ColumnCount; c++)
                    {
                        Cell currentCell = sheet.Cells[r, c];
                        string cellValue = currentCell.StringValue;

                        // Skip empty cells
                        if (string.IsNullOrEmpty(cellValue))
                            continue;

                        // If this value has already been handled, it means this cell is a duplicate
                        if (processedValues.Contains(cellValue))
                        {
                            // Replace duplicate entry with an empty string
                            currentCell.PutValue(string.Empty);
                            continue;
                        }

                        // First occurrence – add to the processed set
                        processedValues.Add(cellValue);

                        // Use Find with the same FindOptions to locate any further duplicates
                        Cell previousFound = null;
                        Cell duplicate = sheet.Cells.Find(cellValue, previousFound, findOptions);
                        bool firstMatch = true; // the first match will be the current cell itself

                        while (duplicate != null)
                        {
                            if (firstMatch && duplicate.Row == r && duplicate.Column == c)
                            {
                                // This is the original occurrence; keep it
                                firstMatch = false;
                            }
                            else
                            {
                                // Clear the duplicate cell
                                duplicate.PutValue(string.Empty);
                            }

                            // Continue searching from the last found cell
                            previousFound = duplicate;
                            duplicate = sheet.Cells.Find(cellValue, previousFound, findOptions);
                        }
                    }
                }

                // Optionally, save the workbook to verify results (file path can be adjusted)
                // workbook.Save("Result.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the console application
        public static void Main()
        {
            Run();
        }
    }
}
