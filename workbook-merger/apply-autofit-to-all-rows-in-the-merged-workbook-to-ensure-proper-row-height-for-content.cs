// Title: Auto‑fit row heights in all worksheets of a merged Excel file with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells auto fit rows after merging multiple Excel workbooks | How to adjust row height for every sheet in a merged .xlsx using Aspose.Cells | Apply AutoFitRows to all worksheets in a combined workbook programmatically in .NET | Set row height automatically based on content in a merged Excel workbook with Aspose.Cells C# | Iterate worksheets and auto‑size rows in a merged workbook using Aspose.Cells
// Tags: AutoFitRows on merged workbook Aspose.Cells | adjust row height after workbook merge .NET | iterate worksheets for row auto sizing C# | use MaxDataRow with AutoFitRows Aspose.Cells | save workbook with updated row heights Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads a previously merged Excel file, iterates each worksheet, determines the last data row with MaxDataRow, applies AutoFitRows to rows 0 through that row, and saves the workbook with correctly sized row heights, handling missing files and exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "merged.xlsx";
        const string outputPath = "merged_autofit.xlsx";

        try
        {
            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the merged workbook
            Workbook workbook = new Workbook(inputPath);

            // Apply AutoFit to all rows in each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                int lastDataRow = sheet.Cells.MaxDataRow;
                if (lastDataRow >= 0)
                {
                    sheet.AutoFitRows(0, lastDataRow);
                }
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
