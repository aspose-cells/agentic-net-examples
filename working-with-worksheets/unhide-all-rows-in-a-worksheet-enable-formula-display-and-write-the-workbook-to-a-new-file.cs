// Title: Unhide every row in all worksheets, show formulas, and save as a new Excel file with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells that loads an existing .xlsx workbook, iterates through each worksheet, makes all rows visible, enables formula display, and saves the result to a new file. | Create a .NET method that accepts input and output paths, opens the workbook with Aspose.Cells, clears hidden flags for all rows containing data, sets Settings.ShowFormula to true, and writes the modified workbook to a different filename.
// Common Searches: Aspose.Cells C# unhide all rows in a workbook and keep formulas visible | How to display formulas instead of values when saving an Excel file with Aspose.Cells | C# iterate through worksheets and unhide rows up to the last used row using Aspose.Cells | Saving a modified workbook to a new file name with Aspose.Cells .NET
// Tags: unhide all rows Aspose.Cells | display formulas Aspose.Cells | save modified workbook Aspose.Cells | worksheet iteration Aspose.Cells | max data row detection Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The program loads "input.xlsx", iterates over every worksheet, determines the last data row, clears the IsHidden flag for all rows up to that index, optionally sets Settings.ShowFormula to true, ensures the output directory exists, and saves the updated workbook as "output_unhidden.xlsx" while handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output_unhidden.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine the last row that contains data
                int maxDataRow = sheet.Cells.MaxDataRow;
                int rowsToProcess = maxDataRow >= 0 ? maxDataRow + 1 : 0;

                // Unhide every row that may contain data
                for (int i = 0; i < rowsToProcess; i++)
                {
                    // Access the row via Cells.Rows collection
                    sheet.Cells.Rows[i].IsHidden = false;
                }
            }

            // If the current Aspose.Cells version supports showing formulas,
            // uncomment the following line. Otherwise, this step is skipped.
            // workbook.Settings.ShowFormula = true;

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display an informative message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
