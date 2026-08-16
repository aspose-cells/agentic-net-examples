// Title: Replace DATEVALUE with native date literals in Excel using Aspose.Cells for .NET (C#)
// Description: Loads a workbook, scans every worksheet for cells that contain the DATEVALUE function, parses the date string argument, converts it to an Excel serial number (honoring the 1904 date system), writes the serial value back, applies a built‑in date format, and saves the file. Includes error handling for malformed formulas and unparsable dates.
// Keywords: Aspose.Cells | C# | DATEVALUE replacement | Excel date serial | 1904 date system | formula conversion | native date literal | Excel automation | batch date update | Excel workbook processing
// Common Searches: how to remove DATEVALUE function with Aspose.Cells | convert DATEVALUE to serial date in .NET | replace Excel DATEVALUE formulas programmatically | Aspose.Cells change DATEVALUE to date literal | C# update Excel dates without DATEVALUE
// Developer Intent: Automatically substitute all DATEVALUE formulas with actual date serial values for clearer, calculation‑free worksheets.
// Use Cases: Migrate legacy spreadsheets that rely on DATEVALUE to static dates before sharing with users who lack formula support. | Batch‑process large workbooks to improve performance by eliminating volatile DATEVALUE calls. | Preserve the original 1904/1900 date system while converting date strings to native Excel dates. | Generate a report of cells where the DATEVALUE argument could not be parsed.
// AI Prompts: Create C# code with Aspose.Cells that finds DATEVALUE formulas, parses the argument, converts it to a serial date respecting the workbook's 1904 setting, writes the value back, and applies a standard date format. | Write a method that scans an Excel file, replaces DATEVALUE calls with literal dates, skips malformed formulas, logs parsing failures, and saves the updated workbook.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads a workbook, scans every worksheet for cells that contain the DATEVALUE function, parses the date string argument, converts it to an Excel serial number (honoring the 1904 date system), writes the serial value back, applies a built‑in date format, and saves the file. Includes error handling for malformed formulas and unparsable dates.
    public class ReplaceDateValueFunction
    {
        public static void Run()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                bool use1904 = workbook.Settings.Date1904;

                // Iterate through worksheets and cells
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (Cell cell in sheet.Cells)
                    {
                        // Process cells containing a DATEVALUE formula
                        if (cell.IsFormula && cell.Formula.IndexOf("DATEVALUE", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            int startIdx = cell.Formula.IndexOf('(');
                            int endIdx = cell.Formula.LastIndexOf(')');
                            if (startIdx < 0 || endIdx < 0 || endIdx <= startIdx + 1)
                                continue; // malformed formula

                            string argument = cell.Formula.Substring(startIdx + 1, endIdx - startIdx - 1).Trim();

                            // Strip surrounding quotes
                            if (argument.StartsWith("\"") && argument.EndsWith("\"") && argument.Length >= 2)
                                argument = argument.Substring(1, argument.Length - 2);

                            // Parse the date string
                            if (DateTime.TryParse(argument, out DateTime dt))
                            {
                                double serial = CellsHelper.GetDoubleFromDateTime(dt, use1904);
                                cell.PutValue(serial);

                                // Apply a built‑in date format for readability
                                Style style = cell.GetStyle();
                                style.Number = 14;
                                cell.SetStyle(style);
                            }
                        }
                    }
                }

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

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            ReplaceDateValueFunction.Run();
        }
    }
}
