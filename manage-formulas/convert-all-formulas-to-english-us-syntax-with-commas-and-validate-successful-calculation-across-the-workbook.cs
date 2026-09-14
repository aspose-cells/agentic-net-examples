// Title: Convert Excel workbook formulas to US English comma syntax and verify calculation results with Aspose.Cells for .NET
// AI Prompts: Load an .xlsx file, set Workbook.Settings.CultureInfo to en-US, recalculate all formulas, and list the addresses of any cells that return error values. | Change the formula locale to US English, execute a full workbook calculation, scan each worksheet for cells whose value starts with '#', and save the updated workbook to a new file.
// Common Searches: Aspose.Cells how to change formula locale to en-US and recalculate workbook | C# detect formula errors after recalculating Excel file with Aspose.Cells | Convert Excel formulas to use commas instead of semicolons using Aspose.Cells .NET | Validate that all formulas in an Excel workbook calculate without errors in C#
// Tags: configure workbook cultureinfo en-us Aspose.Cells | run full formula calculation Aspose.Cells | identify cells with # error values Aspose.Cells | convert formula separators to commas Aspose.Cells | save workbook after formula check Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The example loads an input.xlsx workbook, sets its CultureInfo to en-US to enforce comma‑separated formulas, recalculates every formula, scans all cells for error values (strings beginning with '#'), reports any issues, and saves the processed workbook as output.xlsx.
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
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Set formula locale to English (US) with comma separators
            workbook.Settings.CultureInfo = new CultureInfo("en-US");

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Validate that all formulas calculated without errors
            bool hasError = false;
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                foreach (Cell cell in cells)
                {
                    // Check if the cell contains a formula
                    if (cell.IsFormula)
                    {
                        // After calculation, an error value is represented as a string starting with '#'
                        if (cell.Value is string s && s.StartsWith("#"))
                        {
                            Console.WriteLine($"Error in sheet '{sheet.Name}' cell {cell.Name}");
                            hasError = true;
                        }
                    }
                }
            }

            if (!hasError)
            {
                Console.WriteLine("All formulas calculated successfully.");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
