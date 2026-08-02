// Title: C# – Retrieve Worksheet by Name, Show Formulas, and Save Workbook with Aspose.Cells
// Description: Loads an Excel file, obtains a worksheet by its name (or falls back to the first sheet), enables the ShowFormulas flag so formulas are displayed instead of values, prints the formula in cell A1, and saves the updated workbook.
// Keywords: Aspose.Cells C# retrieve worksheet by name | Aspose.Cells ShowFormulas property | display formulas instead of values Aspose.Cells | save workbook after changing display settings | fallback to first worksheet Aspose.Cells | load Excel file Aspose.Cells .NET | worksheet.ShowFormulas example
// Common Searches: Aspose.Cells get worksheet by name C# | How to display formulas in Excel using Aspose.Cells | ShowFormulas property example | Save workbook after enabling ShowFormulas | What does worksheet.ShowFormulas do | C# code to fallback to first sheet if name not found Aspose.Cells
// Developer Intent: Load an existing workbook, locate a specific worksheet (or default to the first), turn on formula display, and write the modified file.
// Use Cases: Open an Excel file and programmatically select a sheet named "MySheet"; if the sheet does not exist, automatically use the first sheet. | Enable worksheet.ShowFormulas to view all cell formulas instead of calculated results. | Write the modified workbook to a new file while preserving original data. | Log or display the formula text of a particular cell (e.g., A1) after enabling formula view.
// AI Prompts: Generate C# code using Aspose.Cells that loads a workbook, retrieves a worksheet by name with a fallback to the first sheet, sets ShowFormulas = true, and saves the workbook. | Explain how the ShowFormulas property works in Aspose.Cells and when it should be used. | Provide best‑practice error handling for loading a workbook and accessing a worksheet by name in Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsShowFormulasDemo
{
    // Loads an Excel file, obtains a worksheet by its name (or falls back to the first sheet), enables the ShowFormulas flag so formulas are displayed instead of values, prints the formula in cell A1, and saves the updated workbook.
    class Program
    {
        static void Main()
        {
            // Paths for input and output workbooks
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Desired worksheet name
                string worksheetName = "MySheet";

                // Retrieve the worksheet; fall back to the first sheet if not found
                Worksheet worksheet = workbook.Worksheets[worksheetName];
                if (worksheet == null)
                {
                    Console.WriteLine($"Worksheet \"{worksheetName}\" not found. Using the first worksheet instead.");
                    worksheet = workbook.Worksheets[0];
                }

                // Enable formula display for the selected worksheet
                worksheet.ShowFormulas = true;

                // Display the content of cell A1 (will show the formula text if any)
                Console.WriteLine("Cell A1 displayed as: " + worksheet.Cells["A1"].StringValue);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
