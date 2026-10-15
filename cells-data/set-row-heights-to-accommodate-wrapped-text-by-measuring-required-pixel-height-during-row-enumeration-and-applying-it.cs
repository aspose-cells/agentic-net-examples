// Title: How to auto‑fit Excel row heights for wrapped text using Aspose.Cells in C#
// AI Prompts: Write C# code with Aspose.Cells that loops through each worksheet row, detects cells with text wrapping enabled, and calls AutoFitRow only for those rows. | Create a C# method that determines the used range of a workbook, checks each row for wrapped cells, adjusts the row height automatically, and saves the workbook. | Extend the example to also recalculate column widths after adjusting row heights for wrapped text in an Aspose.Cells workbook.
// Common Searches: aspocells c# auto fit row height when text is wrapped | adjust Excel row height based on wrapped cell content using Aspose.Cells | c# iterate worksheet rows and apply AutoFitRow only for wrapped cells | measure required row height for wrapped text Aspose.Cells .NET | how to programmatically auto‑fit rows with text wrap in Aspose.Cells
// Tags: auto-fit row height Aspose.Cells C# | detect wrapped text cells Aspose.Cells | conditional AutoFitRow based on text wrap | row height adjustment for wrapped text .NET | Excel row height auto‑fit using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, iterates through each used row, checks if any cell in the row has text wrapping enabled, and calls AutoFitRow to automatically adjust the row height before saving the updated file.
class RowHeightAdjuster
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the used range
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            // Adjust row heights for rows that contain wrapped text
            for (int row = 0; row <= maxRow; row++)
            {
                bool hasWrapped = false;

                // Check each cell in the current row for text wrapping
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    Style style = cell.GetStyle();

                    if (style.IsTextWrapped && !string.IsNullOrEmpty(cell.StringValue))
                    {
                        hasWrapped = true;
                        break;
                    }
                }

                // Auto‑fit the row height if any wrapped cell is present
                if (hasWrapped)
                {
                    sheet.AutoFitRow(row);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
