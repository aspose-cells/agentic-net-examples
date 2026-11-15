// Title: AutoFit all columns in each worksheet of a merged Excel workbook using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells how to auto‑fit columns after merging multiple Excel files | apply column AutoFit to every sheet in a combined workbook with Aspose.Cells .NET | auto size all columns in merged.xlsx using Aspose.Cells API | Aspose.Cells .NET iterate worksheets and AutoFitColumn for each column | adjust column widths programmatically after workbook merge in C#
// Tags: Aspose.Cells AutoFitColumn across worksheets | auto fit columns merged workbook .NET | iterate worksheets adjust column width C# | column width optimization after Excel merge | programmatic column autosizing Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads a merged Excel file (merged.xlsx) with Aspose.Cells, loops through every worksheet, determines the last used column, calls AutoFitColumn for each column up to that index, and saves the adjusted workbook as merged_autofit.xlsx, handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        const string inputFile = "merged.xlsx";
        const string outputFile = "merged_autofit.xlsx";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file '{inputFile}' not found.");
                return;
            }

            // Load the merged workbook
            Workbook workbook = new Workbook(inputFile);

            // Apply AutoFit to all columns in each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine the last used column (zero‑based)
                int lastColumn = sheet.Cells.MaxColumn;

                // AutoFit each column up to the last used one
                for (int col = 0; col <= lastColumn; col++)
                {
                    sheet.AutoFitColumn(col);
                }
            }

            // Save the updated workbook
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully as '{outputFile}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
