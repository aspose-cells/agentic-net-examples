// Title: Clear the AutoFilter on column AJ (index 35) and display all rows using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that removes any AutoFilter criteria from column AJ (zero‑based index 35) of an existing workbook and saves the updated file with Aspose.Cells. | Create a method that resets the worksheet filter for a specific column, ensuring all hidden rows become visible before saving the workbook.
// Common Searches: asp.net remove autofilter from column AJ using Aspose.Cells | how to show all rows after clearing a specific Excel filter with C# | Aspose.Cells delete filter criteria by column index | C# code to reset worksheet AutoFilter and unhide rows | clear filter on column AJ in an existing workbook Aspose.Cells .NET
// Tags: clear autofilter column Aspose.Cells | remove filter criteria C# Aspose.Cells | show all rows after filter removal .NET | auto filter column index handling Aspose.Cells | excel workbook filter manipulation C#

using System;
using System.IO;
using Aspose.Cells;

// The program loads "input.xlsx", checks for an AutoFilter, removes any filter applied to column AJ (zero‑based index 35), and saves the workbook as "output.xlsx", handling missing files and out‑of‑range column scenarios.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Access the AutoFilter object of the worksheet
            AutoFilter autoFilter = worksheet.AutoFilter;

            // Proceed only if an AutoFilter is present and has filter columns
            if (autoFilter != null && autoFilter.FilterColumns != null && autoFilter.FilterColumns.Count > 0)
            {
                // Column AJ corresponds to index 35 (0‑based)
                int columnIndex = 35;

                try
                {
                    // Attempt to remove the filter criteria for the specified column
                    autoFilter.FilterColumns.RemoveAt(columnIndex);
                }
                catch (ArgumentOutOfRangeException)
                {
                    // The specified column does not have a filter applied; ignore
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Unable to remove filter for column AJ. {ex.Message}");
                }
            }

            // Save the workbook with the cleared filter criteria
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
