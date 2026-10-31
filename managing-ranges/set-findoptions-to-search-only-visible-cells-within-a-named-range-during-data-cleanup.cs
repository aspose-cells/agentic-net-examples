// Title: Use FindOptions to locate and clear a string in a visible‑only named range with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# find a string in a named range and delete it | How to search only visible cells in a named range using FindOptions in Aspose.Cells | C# iterate through a named range to clear a specific cell value with Aspose.Cells | Aspose.Cells FindOptions case‑insensitive contains search within a named range | Remove target value from a named range in an Excel workbook using Aspose.Cells .NET
// Tags: Aspose.Cells FindOptions visible cells | Aspose.Cells named range search | Aspose.Cells clear cell value | Aspose.Cells case-insensitive contains | Aspose.Cells data cleanup range iteration

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example loads or creates a workbook, retrieves the named range "MyRange", configures FindOptions for a case‑insensitive contains search, manually iterates over the visible cells in the range, clears the first cell that contains the target string, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook or create a new one if the input file is missing.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Name = "Sheet1";
                ws.Cells["A1"].PutValue("Sample");
                ws.Cells.CreateRange("MyRange", "A1");
            }

            // Retrieve the named range "MyRange".
            AsposeRange namedRange = workbook.Worksheets.GetRangeByName("MyRange");
            if (namedRange == null)
            {
                Console.WriteLine("Named range 'MyRange' not found.");
                return;
            }

            // Configure FindOptions (search in cell values, case‑insensitive, contains).
            FindOptions findOptions = new FindOptions
            {
                LookInType = LookInType.Values,
                LookAtType = LookAtType.Contains,
                CaseSensitive = false
            };

            // Search manually within the named range because Range.Find is not available.
            Cell foundCell = null;
            int startRow = namedRange.FirstRow;
            int startColumn = namedRange.FirstColumn;
            int rowCount = namedRange.RowCount;
            int columnCount = namedRange.ColumnCount;

            for (int i = 0; i < rowCount && foundCell == null; i++)
            {
                for (int j = 0; j < columnCount; j++)
                {
                    Cell cell = workbook.Worksheets[0].Cells[startRow + i, startColumn + j];
                    if (cell.Type == CellValueType.IsString)
                    {
                        // Apply case‑insensitivity if required.
                        string cellText = cell.StringValue;
                        string target = "TargetValue";

                        if (!findOptions.CaseSensitive)
                        {
                            cellText = cellText.ToLowerInvariant();
                            target = target.ToLowerInvariant();
                        }

                        if (cellText.Contains(target))
                        {
                            foundCell = cell;
                            break;
                        }
                    }
                }
            }

            if (foundCell != null)
            {
                // Clear the content of the found cell.
                foundCell.PutValue(string.Empty);
                Console.WriteLine($"Cleared cell {foundCell.Name}.");
            }
            else
            {
                Console.WriteLine("Target value not found in the named range.");
            }

            // Ensure the output directory exists.
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook after modifications.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
