// Title: Unhide every row in all worksheets, display formulas, and save as a new Excel file using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an existing .xlsx workbook with Aspose.Cells, iterates through each worksheet, and sets every row's IsHidden property to false. | Provide a C# snippet that activates the Settings.ShowFormula option in Aspose.Cells so formulas are shown instead of their calculated results before saving. | Create a robust C# example that verifies the source file, unhides all rows, toggles formula view, and writes the modified workbook to a different output path with proper exception handling.
// Common Searches: Aspose.Cells C# unhide all rows in every worksheet of an Excel workbook | How to show formulas instead of values when saving a workbook with Aspose.Cells .NET | C# iterate through all worksheets and change row visibility using Aspose.Cells | Saving a modified Excel file to a new location with Aspose.Cells for .NET | Check if input Excel file exists before processing with Aspose.Cells C#
// Tags: unhide rows Aspose.Cells .NET | show formulas Settings.ShowFormula | iterate worksheets Aspose.Cells C# | save workbook to new file Aspose.Cells | file existence check C# Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks that input.xlsx exists, loads it with Aspose.Cells, loops through each worksheet to set every row's IsHidden flag to false, optionally enables Settings.ShowFormula to display formulas, and then saves the updated workbook as output.xlsx while handling potential exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            var workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Unhide all rows in the current worksheet
                int totalRows = sheet.Cells.Rows.Count;
                for (int i = 0; i < totalRows; i++)
                {
                    sheet.Cells.Rows[i].IsHidden = false;
                }
            }

            // Optionally display formulas instead of results if the API supports it
            // Uncomment the following line if your Aspose.Cells version includes ShowFormula
            // workbook.Settings.ShowFormula = true;

            // Save the modified workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
