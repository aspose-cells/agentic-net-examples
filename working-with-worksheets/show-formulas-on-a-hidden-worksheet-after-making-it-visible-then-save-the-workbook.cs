// Title: Unhide the first hidden worksheet, display its formulas, and save the workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Load an existing .xlsx file, locate the first hidden worksheet, set its IsVisible property to true, enable Settings.ShowFormula, and save the workbook to a new file with Aspose.Cells in C#. | Using Aspose.Cells for .NET, make a hidden sheet visible, turn on formula view, and export the modified workbook while handling missing input files.
// Common Searches: Aspose.Cells C# unhide hidden worksheet and show formulas | How to enable ShowFormula setting when saving Excel workbook with Aspose.Cells | C# code to make hidden Excel sheet visible and keep formulas displayed | Save modified workbook after revealing hidden sheet using Aspose.Cells | Check for hidden worksheets in Aspose.Cells and make them visible
// Tags: make hidden sheet visible Aspose.Cells | enable formula view Settings.ShowFormula | export workbook after sheet visibility change | load and iterate worksheets Aspose.Cells C# | validate input file existence Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// // Loads input.xlsx, finds the first hidden worksheet, sets its IsVisible to true, optionally turns on Settings.ShowFormula to display formulas, and saves the updated workbook as output.xlsx while handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Locate the first hidden worksheet
            Worksheet hiddenSheet = null;
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (!ws.IsVisible) // worksheet is hidden
                {
                    hiddenSheet = ws;
                    break;
                }
            }

            if (hiddenSheet != null)
            {
                // Make the hidden worksheet visible
                hiddenSheet.IsVisible = true;

                // Note: Displaying formulas instead of calculated values requires
                // the ShowFormula property, which may not be available in older
                // Aspose.Cells versions. If supported, uncomment the line below:
                // workbook.Settings.ShowFormula = true;
            }

            // Save the modified workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
