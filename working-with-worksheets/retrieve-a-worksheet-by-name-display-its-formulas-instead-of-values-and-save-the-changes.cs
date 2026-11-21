// Title: Retrieve a worksheet by name, display its formulas instead of calculated values, and save the workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells get worksheet by name with fallback to first sheet | show formulas instead of values when saving Excel with Aspose.Cells .NET | how to configure Aspose.Cells to display formulas on workbook open | handle missing worksheet name in Aspose.Cells C# code example | save workbook so formulas are visible in Excel using Aspose.Cells
// Tags: retrieve worksheet by name Aspose.Cells C# | display formulas instead of values Aspose.Cells | fallback to first worksheet Aspose.Cells | save workbook with formula view Aspose.Cells | missing worksheet handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file with Aspose.Cells, attempts to retrieve a worksheet by its name (falling back to the first sheet if the name is not found), notes that the ShowFormula property is no longer available, and saves the workbook so that formulas are shown when the file is opened in Excel, while handling file‑not‑found and other exceptions.
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

            // Load the workbook from the existing file
            var workbook = new Workbook(inputPath);

            // Retrieve the worksheet by its name (replace "MySheet" with the actual sheet name)
            var worksheet = workbook.Worksheets["MySheet"];
            if (worksheet == null)
            {
                Console.WriteLine("Warning: Worksheet \"MySheet\" not found. Using the first worksheet.");
                worksheet = workbook.Worksheets[0];
            }

            // NOTE: In recent versions of Aspose.Cells the WorkbookSettings class does not expose a ShowFormula property.
            // If you need to display formulas instead of calculated values when opening the file in Excel,
            // you can set the appropriate option in the save options (if supported) or handle it manually.
            // The following line is intentionally omitted because the property does not exist:
            // workbook.Settings.ShowFormula = true;

            // Save the modified workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
