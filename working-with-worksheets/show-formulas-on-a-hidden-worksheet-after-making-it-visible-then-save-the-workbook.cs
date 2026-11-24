// Title: Unhide hidden worksheets and enable formula display before saving the workbook using Aspose.Cells for .NET
// Common Searches: how to make hidden Excel sheets visible with Aspose.Cells C# | display cell formulas on a worksheet after unhiding it using Aspose.Cells | Aspose.Cells C# set ShowFormulas for hidden worksheets before saving | iterate through worksheets to unhide and show formulas in a .NET application | save workbook after changing sheet visibility and formula view with Aspose.Cells
// Tags: unhide hidden sheet Aspose.Cells | show formulas on worksheet Aspose.Cells | ShowFormulas property C# | iterate worksheets visibility Aspose.Cells | save workbook after visibility change .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Loads an Excel file, iterates over all worksheets, makes any hidden sheet visible, enables formula display with ShowFormulas, and saves the updated workbook to a new file.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook from the input file
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets to find hidden ones
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    if (!sheet.IsVisible)
                    {
                        // Make hidden worksheet visible
                        sheet.IsVisible = true;

                        // Enable formula display for this worksheet
                        try
                        {
                            sheet.ShowFormulas = true;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Unable to set ShowFormulas for sheet '{sheet.Name}': {ex.Message}");
                        }
                    }
                }

                // Save the modified workbook to the output file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
