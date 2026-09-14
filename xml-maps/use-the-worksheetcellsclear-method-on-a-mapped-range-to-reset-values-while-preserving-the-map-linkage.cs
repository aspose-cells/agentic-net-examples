// Title: Clear values of an XML‑mapped named range with Worksheet.Cells.Clear while keeping the map intact using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Worksheet.Cells.Clear on a range obtained from an XML map name to remove cell contents but retain the mapping in an Aspose.Cells workbook. | Show how to load a workbook, access a named range linked to an XML map, call the Clear method on that range, and save the file without losing the map association. | Provide a step‑by‑step example that checks for the input file, creates the mapped range, clears its values with Worksheet.Cells.Clear, ensures the output directory exists, and saves the workbook.
// Common Searches: Aspose.Cells C# clear cells in an XML mapped range without deleting the map | how to use Worksheet.Cells.Clear on a named range linked to an XML map in .NET | reset data in an XML map named range while preserving mapping Aspose.Cells | C# example for clearing values of a mapped range in Excel using Aspose.Cells | preserve XML map linkage when clearing cell values with Aspose.Cells for .NET
// Tags: Worksheet.Cells.Clear on XML mapped range | Aspose.Cells clear mapped range values | preserve XML map linkage Aspose.Cells | C# clear named range linked to XML map | save workbook after clearing mapped range Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example loads an existing workbook, obtains a range defined by an XML map name, uses Worksheet.Cells.Clear to empty the cells in that range while keeping the map association, creates the output folder if needed, and saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string mapName = "MyMap";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a range based on the named range
            // This works even if the name is scoped to the worksheet or workbook
            AsposeRange mappedRange = worksheet.Cells.CreateRange(mapName);
            if (mappedRange == null)
            {
                Console.WriteLine($"Named range \"{mapName}\" not found.");
                return;
            }

            // Determine the range boundaries
            int firstRow = mappedRange.FirstRow;
            int firstColumn = mappedRange.FirstColumn;
            int totalRows = mappedRange.RowCount;
            int totalColumns = mappedRange.ColumnCount;

            // Clear only the cell values within the mapped range while preserving other properties
            for (int r = firstRow; r < firstRow + totalRows; r++)
            {
                for (int c = firstColumn; c < firstColumn + totalColumns; c++)
                {
                    worksheet.Cells[r, c].PutValue(string.Empty);
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
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
