// Title: Rename each Excel worksheet to "Sheet_{row+col}" based on MaxDataRow and MaxDataColumn using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells rename worksheet to Sheet_ based on MaxDataRow and MaxDataColumn sum | How to generate unique worksheet names from data dimensions with Aspose.Cells .NET | Programmatically set Excel sheet name to row plus column count using Aspose.Cells | Iterate worksheets and rename them using MaxDataRow and MaxDataColumn in C# | Avoid duplicate sheet names when renaming worksheets by data size in Aspose.Cells
// Tags: worksheet renaming using MaxDataRow Aspose.Cells | unique sheet name generation Aspose.Cells .NET | calculate worksheet data size C# Aspose.Cells | duplicate sheet name handling Aspose.Cells | auto rename Excel sheets by row column sum Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace WorksheetRenamer
{
    // The example loads an Excel workbook with Aspose.Cells, loops through each worksheet, obtains its MaxDataRow and MaxDataColumn (treating -1 as 0), computes their sum, builds a name like "Sheet_{sum}" while ensuring the name is unique within the workbook, assigns the new name to the worksheet, and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths.
            // Adjust these paths as needed for your environment.
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists.
            if (!File.Exists(inputPath))
            {
                Console.Error.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                return;
            }

            try
            {
                // Load the workbook.
                Workbook workbook = new Workbook(inputPath);
                WorksheetCollection worksheets = workbook.Worksheets;

                Console.WriteLine($"Workbook loaded. Contains {worksheets.Count} worksheet(s).");

                // Iterate through each worksheet and rename it.
                for (int i = 0; i < worksheets.Count; i++)
                {
                    Worksheet sheet = worksheets[i];
                    // MaxDataRow and MaxDataColumn return -1 when there is no data.
                    int maxRow = sheet.Cells.MaxDataRow;
                    int maxCol = sheet.Cells.MaxDataColumn;

                    // Convert -1 to 0 for a more intuitive sum.
                    int effectiveRow = maxRow >= 0 ? maxRow : 0;
                    int effectiveCol = maxCol >= 0 ? maxCol : 0;

                    int sum = effectiveRow + effectiveCol;

                    // Build a new name. Ensure the name is unique within the workbook.
                    string baseName = $"Sheet_{sum}";
                    string newName = baseName;
                    int duplicateCounter = 1;
                    while (IsWorksheetNameTaken(worksheets, newName, i))
                    {
                        newName = $"{baseName}_{duplicateCounter}";
                        duplicateCounter++;
                    }

                    Console.WriteLine($"Renaming worksheet \"{sheet.Name}\" to \"{newName}\" (MaxRow={maxRow}, MaxCol={maxCol}, Sum={sum}).");
                    sheet.Name = newName;
                }

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("An unexpected error occurred:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }

        /// <summary>
        /// Checks whether a worksheet name is already used in the collection,
        /// excluding the worksheet at the specified index (the one currently being renamed).
        /// </summary>
        /// <param name="worksheets">The collection of worksheets.</param>
        /// <param name="name">The name to test.</param>
        /// <param name="excludeIndex">Index of the worksheet being renamed.</param>
        /// <returns>True if the name exists elsewhere; otherwise false.</returns>
        private static bool IsWorksheetNameTaken(WorksheetCollection worksheets, string name, int excludeIndex)
        {
            for (int i = 0; i < worksheets.Count; i++)
            {
                if (i == excludeIndex) continue;
                if (string.Equals(worksheets[i].Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }
    }
}
