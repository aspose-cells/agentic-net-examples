// Title: Retrieve a worksheet by name, display its formulas, and save the workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an existing Excel file using Aspose.Cells, selects a worksheet by its name (or the first sheet if the name is missing), sets ShowFormulas = true, and saves the result to a new file. | Generate a try‑catch block that checks whether the input Excel file exists, loads the workbook, toggles the ShowFormulas property on a specific worksheet, and handles any exceptions while saving the workbook. | Create a reusable C# method that accepts inputPath, outputPath, and sheetName, loads the workbook with Aspose.Cells, ensures the worksheet is present, enables formula view, and writes the modified workbook to the output path.
// Common Searches: Aspose.Cells C# show formulas for a specific worksheet and save workbook | How to get worksheet by name with fallback to first sheet using Aspose.Cells .NET | Set ShowFormulas property in Aspose.Cells and export to a new Excel file | C# verify Excel file exists before loading with Aspose.Cells | Display formulas instead of values in an Aspose.Cells workbook
// Tags: get worksheet by name Aspose.Cells | enable ShowFormulas property C# | fallback to first worksheet Aspose.Cells | save modified workbook Aspose.Cells | verify input Excel file existence C#

using System;
using System.IO;
using Aspose.Cells;

// // Loads input.xlsx, obtains the worksheet named "MySheet" (or the first sheet if not found), enables ShowFormulas to display formulas instead of calculated values, and saves the modified workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string sheetName = "MySheet";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the worksheet by name; fallback to the first worksheet if not found
            Worksheet sheet = workbook.Worksheets[sheetName];
            if (sheet == null)
            {
                Console.WriteLine($"Worksheet \"{sheetName}\" not found. Using the first worksheet.");
                sheet = workbook.Worksheets[0];
            }

            // Display formulas instead of calculated values
            sheet.ShowFormulas = true;

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
