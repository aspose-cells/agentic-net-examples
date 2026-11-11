// Title: How to use Aspose.Cells FindOptions for a case‑sensitive search in column E (E1:E100) with C#
// Common Searches: Aspose.Cells C# find text case sensitive within a specific column range | C# FindOptions case sensitive search limited to rows 1 to 100 in Excel using Aspose.Cells | Search for a value in column E (E1:E100) with case sensitivity using Aspose.Cells API | How to restrict Aspose.Cells Find method to a defined range in C# | Aspose.Cells FindOptions LookInType.Values case sensitive example
// Tags: Aspose.Cells FindOptions case-sensitive search | C# search specific range E1:E100 Aspose.Cells | Excel column E text lookup Aspose.Cells | limit Find method rows 1-100 Aspose.Cells | LookInType.Values usage Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a workbook, sets the start cell to E1, configures FindOptions with CaseSensitive = true and LookInType = Values, searches for the string "Target" starting from that cell, checks that the found cell is within rows 1‑100 of column E, outputs the result, and saves the workbook as output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the start cell of the search range (E1)
                Cell startCell = worksheet.Cells[0, 4]; // Row 0 (E1), Column 4 (E)

                // Configure FindOptions for a case‑sensitive search
                FindOptions findOptions = new FindOptions
                {
                    CaseSensitive = true,
                    LookInType = LookInType.Values
                };

                // Search for the text "Target" starting from the defined cell
                Cell foundCell = worksheet.Cells.Find("Target", startCell, findOptions);

                // Verify that the found cell is within rows 1‑100 (indices 0‑99)
                if (foundCell != null && foundCell.Row <= 99 && foundCell.Column == 4)
                {
                    Console.WriteLine($"Found at: {foundCell.Name}");
                }
                else
                {
                    Console.WriteLine("Target not found in the specified range.");
                }

                // Save the workbook
                string outputPath = "output.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
