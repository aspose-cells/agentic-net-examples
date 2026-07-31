// Title: Rename Excel worksheets by sum of populated rows and columns using Aspose.Cells for .NET
// Description: A C# utility that loads or creates a workbook, iterates through all worksheets, calculates (MaxDataRow+1)+(MaxDataColumn+1), generates a safe sheet name with CellsHelper.CreateSafeSheetName, resolves duplicates, renames each sheet, and saves the result.
// Keywords: Aspose.Cells | C# rename worksheet | MaxDataRow | MaxDataColumn | Excel sheet naming | CreateSafeSheetName | unique worksheet name | batch rename Excel sheets | Excel automation .NET | worksheet rename example
// Common Searches: rename Excel sheets based on data size Aspose.Cells | C# sum of rows and columns for worksheet name | ensure unique sheet names with Aspose.Cells | CellsHelper.CreateSafeSheetName usage | how to rename all worksheets in a workbook programmatically
// Developer Intent: Rename each worksheet to a unique name that reflects the total count of populated rows plus columns.
// Use Cases: Quickly identify sheet size by embedding row‑plus‑column count in the sheet name. | Process large workbooks in batch, guaranteeing Excel‑compliant and non‑duplicate sheet names. | Create a new workbook when the source file is missing and apply the same naming logic to the default sheet. | Integrate into automated reporting pipelines where sheet names must convey data volume.
// AI Prompts: Generate C# code with Aspose.Cells that renames every worksheet to "Sum_{rowCount+colCount}" while ensuring the name is Excel‑safe and unique. | Explain the role of CellsHelper.CreateSafeSheetName and show how to handle duplicate worksheet names during renaming. | Provide a helper method to detect existing sheet names in a Workbook and append an incremental suffix to avoid conflicts.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsRenameWorksheets
{
    // A C# utility that loads or creates a workbook, iterates through all worksheets, calculates (MaxDataRow+1)+(MaxDataColumn+1), generates a safe sheet name with CellsHelper.CreateSafeSheetName, resolves duplicates, renames each sheet, and saves the result.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new one.
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook(); // creates a workbook with a default sheet
                }

                // Iterate through all worksheets in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // MaxDataRow and MaxDataColumn are zero‑based indexes of the last cell that contains data.
                    // Adding 1 to each gives the count of rows/columns with data.
                    int maxRow = sheet.Cells.MaxDataRow + 1;
                    int maxCol = sheet.Cells.MaxDataColumn + 1;

                    // Calculate the sum for naming
                    int sum = maxRow + maxCol;

                    // Build a provisional name
                    string provisionalName = $"Sum_{sum}";

                    // Ensure the name complies with Excel rules (max 31 chars, no illegal characters)
                    string safeName = CellsHelper.CreateSafeSheetName(provisionalName);

                    // Ensure uniqueness within the workbook
                    string finalName = safeName;
                    int duplicateCounter = 1;
                    while (IsWorksheetNameExists(workbook.Worksheets, finalName, sheet.Index))
                    {
                        finalName = $"{safeName}_{duplicateCounter}";
                        finalName = CellsHelper.CreateSafeSheetName(finalName);
                        duplicateCounter++;
                    }

                    // Rename the worksheet
                    sheet.Name = finalName;
                }

                // Save the modified workbook
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper method to check if a name already exists in the collection,
        // excluding the worksheet at the specified index (the one being renamed).
        static bool IsWorksheetNameExists(WorksheetCollection sheets, string name, int excludeIndex)
        {
            foreach (Worksheet ws in sheets)
            {
                if (ws.Index != excludeIndex && string.Equals(ws.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }
    }
}
