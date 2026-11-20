// Title: How to auto‑fit Excel column widths in C# with Aspose.Cells based on the longest cell content
// Common Searches: C# Aspose.Cells auto fit column width to longest cell value | adjust Excel column width programmatically using Aspose.Cells .NET | set column width in Excel workbook according to cell text length C# | dynamic column sizing in Aspose.Cells based on content measurement | auto‑size Excel columns after loading a workbook with Aspose.Cells
// Tags: Aspose.Cells auto fit column width | C# adjust Excel column width by content | dynamic column sizing Aspose.Cells .NET | measure cell text pixel width Excel | programmatic column width optimization workbook

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, determines the used range, iterates over each column, and calls AutoFitColumn to automatically size columns according to the longest cell content before saving the updated file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Determine the used range
            int maxRow = cells.MaxDataRow;      // zero‑based index of the last row with data
            int maxCol = cells.MaxDataColumn;   // zero‑based index of the last column with data

            // Auto‑fit each column based on its content
            for (int col = 0; col <= maxCol; col++)
            {
                try
                {
                    // Adjust column width to fit the longest cell in the column
                    sheet.AutoFitColumn(col);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to autofit column {col}: {ex.Message}");
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
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
