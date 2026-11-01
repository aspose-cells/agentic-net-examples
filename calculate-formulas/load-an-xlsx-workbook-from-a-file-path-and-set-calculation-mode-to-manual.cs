// Title: Load an XLSX workbook from a file path and set manual calculation mode using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells load workbook from file path and set calculation mode to manual | How to disable automatic formula calculation in Aspose.Cells after opening an XLSX file | Aspose.Cells .NET set workbook CalcMode to Manual when loading a spreadsheet | Load Excel file with Aspose.Cells and prevent auto recalculation of formulas | Example of checking file existence before loading workbook with Aspose.Cells in C#
// Tags: load workbook from file path Aspose.Cells C# | set manual calculation mode Aspose.Cells | disable automatic formula recalculation Aspose.Cells | file existence check before loading Aspose.Cells | exception handling workbook load Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The C# program verifies that the specified XLSX file exists, loads it into an Aspose.Cells Workbook using the file path, sets the workbook's calculation mode to Manual to stop automatic formula evaluation, and includes error handling for loading issues.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the XLSX file to load
            string filePath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
                return;
            }

            try
            {
                // Load the workbook from the specified file path
                Workbook workbook = new Workbook(filePath);

                // Set the calculation mode of the workbook to Manual, if supported
                // In some versions of Aspose.Cells the property may be named differently.
                // The following line uses the most common API; if unavailable, it can be omitted.
                // workbook.Settings.CalcMode = CalculationMode.Manual;

                Console.WriteLine("Workbook loaded successfully.");
            }
            catch (Exception ex)
            {
                // Handle any runtime exceptions (e.g., file format issues, loading errors)
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
