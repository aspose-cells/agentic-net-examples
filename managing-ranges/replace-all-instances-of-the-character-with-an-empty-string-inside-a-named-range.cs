// Title: Remove asterisk (*) characters from every cell in a named range with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells replace asterisk in all cells of a named range | How to delete * characters from an Excel named range using Aspose.Cells | Iterate through GetRangeByName and remove specific symbols in C# | Aspose.Cells example to clean string values in a named range | Save workbook after modifying cells in a named range with Aspose.Cells
// Tags: replace asterisk in named range Aspose.Cells | Aspose.Cells GetRangeByName string cleanup | iterate over cells in Excel named range C# | remove specific character from Excel cells Aspose.Cells | save modified workbook Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, obtains a named range called "MyRange" via GetRangeByName, iterates through each cell, removes any asterisk (*) characters from string values, and saves the updated workbook to a new file.
    class ReplaceAsteriskInNamedRange
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string rangeName = "MyRange";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Retrieve the named range using the correct API
                AsposeRange? namedRange = workbook.Worksheets.GetRangeByName(rangeName);

                if (namedRange != null)
                {
                    // Iterate through each cell in the named range
                    foreach (Cell cell in namedRange)
                    {
                        // Process only string cells
                        if (cell.Type == CellValueType.IsString && cell.Value != null)
                        {
                            string original = cell.StringValue;
                            string replaced = original.Replace("*", string.Empty);
                            cell.PutValue(replaced);
                        }
                    }

                    // Ensure the output directory exists
                    string outputDir = Path.GetDirectoryName(outputPath);
                    if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    // Save the modified workbook
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved successfully to {outputPath}");
                }
                else
                {
                    Console.WriteLine($"Named range '{rangeName}' was not found in the workbook.");
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
