// Title: Remove duplicate entries from a named range with FindOptions – Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, defines a named range (A1:A6), and uses Aspose.Cells FindOptions limited to that range to locate and clear duplicate values, preserving only the first occurrence. The cleaned workbook is saved as an XLSX file.
// Keywords: Aspose.Cells duplicate removal | named range FindOptions C# | remove duplicate cells .NET | Aspose.Cells Find method example | clear duplicate entries Excel library
// Common Searches: Aspose.Cells delete duplicate values in named range | FindOptions duplicate search C# | how to remove duplicate cells with Aspose.Cells | C# code to clean named range duplicates | Aspose.Cells de‑duplication example
// Developer Intent: Programmatically eliminate repeated values inside a defined named range, keeping the first occurrence of each entry.
// Use Cases: Sanitize a list of product codes stored in a named range before exporting. | Prepare data for validation by ensuring a named range contains unique items. | Automate cleanup of user‑entered lists (e.g., categories, tags) in Excel worksheets.
// AI Prompts: Write C# code using Aspose.Cells to remove duplicate rows from a multi‑column named range while retaining the first row. | Show how to modify the example to ignore case‑sensitive duplicates when searching with FindOptions. | Explain how to set LookAtType.Partial in FindOptions to delete cells that partially match a given string within a named range.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example creates a workbook, defines a named range (A1:A6), and uses Aspose.Cells FindOptions limited to that range to locate and clear duplicate values, preserving only the first occurrence. The cleaned workbook is saved as an XLSX file.
    public class RemoveDuplicateEntriesInNamedRange
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

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
                int nameIndex = sheet.Workbook.Worksheets.Names.Add("MyRange");
                sheet.Workbook.Worksheets.Names[nameIndex].RefersTo = $"={sheet.Name}!$A$1:$A$6";

                // Retrieve the named range object
                Name namedRange = sheet.Workbook.Worksheets.Names["MyRange"];
                Aspose.Cells.Range range = namedRange.GetRange();

                // Prepare FindOptions with the same search range as the named range
                FindOptions findOptions = new FindOptions
                {
                    LookInType = LookInType.Values
                    // LookAtType defaults to whole content; explicit setting removed to avoid compatibility issues
                };

                CellArea searchArea = new CellArea
                {
                    StartRow = range.FirstRow,
                    StartColumn = range.FirstColumn,
                    EndRow = range.FirstRow + range.RowCount - 1,
                    EndColumn = range.FirstColumn + range.ColumnCount - 1
                };
                findOptions.SetRange(searchArea);

                // Keep track of values already processed to avoid re‑scanning them
                HashSet<string> processedValues = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                // Iterate through each cell in the named range
                for (int row = range.FirstRow; row <= range.FirstRow + range.RowCount - 1; row++)
                {
                    for (int col = range.FirstColumn; col <= range.FirstColumn + range.ColumnCount - 1; col++)
                    {
                        Cell currentCell = sheet.Cells[row, col];
                        string cellValue = currentCell.StringValue;

                        // Skip empty cells
                        if (string.IsNullOrEmpty(cellValue))
                            continue;

                        // If this value has already been processed, it means the first occurrence
                        // was earlier and this cell is a duplicate; clear it.
                        if (processedValues.Contains(cellValue))
                        {
                            currentCell.PutValue(string.Empty);
                            continue;
                        }

                        // First time we see this value – add to processed set
                        processedValues.Add(cellValue);

                        // Find subsequent duplicates within the same range
                        Cell previousFound = currentCell;
                        Cell duplicateCell = sheet.Cells.Find(cellValue, previousFound, findOptions);
                        while (duplicateCell != null)
                        {
                            // Clear the duplicate cell
                            duplicateCell.PutValue(string.Empty);

                            // Continue searching after the cleared cell
                            previousFound = duplicateCell;
                            duplicateCell = sheet.Cells.Find(cellValue, previousFound, findOptions);
                        }
                    }
                }

                // Save the workbook with duplicates removed
                workbook.Save("RemoveDuplicatesInNamedRange.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Run error: {ex.Message}");
                throw; // Re‑throw to be caught by outer handler if needed
            }
        }
    }
}
