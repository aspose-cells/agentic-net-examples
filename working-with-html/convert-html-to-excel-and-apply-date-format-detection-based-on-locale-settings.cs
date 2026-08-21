// Title: C# – Convert HTML to Excel with locale‑aware date detection using Aspose.Cells
// Description: Load an HTML file into an Aspose.Cells workbook, scan the used range, parse string values as dates with the current CultureInfo, replace them with DateTime cells, apply the built‑in short date style (Number 14), and save as XLSX.
// Keywords: Aspose.Cells HTML to Excel | C# locale date parsing | CultureInfo date conversion | Excel short date format index 14 | detect dates in HTML tables | convert HTML tables to XLSX | Aspose.Cells date detection
// Common Searches: Aspose.Cells convert HTML to Excel C# | how to detect dates from HTML using CultureInfo in Aspose.Cells | apply short date format to cells after loading HTML | C# parse locale specific dates in Excel workbook | batch convert HTML reports to XLSX with date handling
// Developer Intent: Transform an HTML document into an Excel workbook while automatically converting locale‑specific date strings into proper DateTime cells with a consistent short‑date display.
// Use Cases: Migrate web‑based reports (HTML tables) to Excel while preserving date semantics across regional settings. | Process legacy system exports that deliver HTML tables, ensuring date columns become true Excel dates. | Automate bulk conversion of multiple HTML files to XLSX with reliable date detection and formatting.
// AI Prompts: Generate C# code that uses Aspose.Cells to load an HTML file, detect date strings based on CultureInfo.CurrentCulture, replace them with DateTime values, set the short date style (Number 14), and save the workbook as XLSX. | Create a reusable method for a Workbook that scans its used range, parses locale‑aware dates from string cells, applies a short date format, and returns the updated workbook.

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// Load an HTML file into an Aspose.Cells workbook, scan the used range, parse string values as dates with the current CultureInfo, replace them with DateTime cells, apply the built‑in short date style (Number 14), and save as XLSX.
class HtmlToExcelConverter
{
    static void Main()
    {
        try
        {
            // Input HTML file and output Excel file paths
            string htmlPath = "input.html";
            string excelPath = "output.xlsx";

            // Verify that the input HTML file exists
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Input file not found: {htmlPath}");
                return;
            }

            // Load the HTML file into a workbook
            Workbook workbook = new Workbook(htmlPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the used range of cells (fully qualified to avoid ambiguity)
            Aspose.Cells.Range usedRange = worksheet.Cells.MaxDisplayRange;

            // Detect date strings based on the current locale and convert them to DateTime values
            foreach (Cell cell in usedRange)
            {
                if (cell.Type == CellValueType.IsString)
                {
                    string cellText = cell.StringValue.Trim();

                    // Try to parse the string as a date using the current culture
                    if (DateTime.TryParse(cellText, CultureInfo.CurrentCulture, DateTimeStyles.None, out DateTime parsedDate))
                    {
                        // Replace the string with a DateTime value
                        cell.PutValue(parsedDate);

                        // Apply a built‑in short date format (index 14) to keep the display consistent
                        Style style = cell.GetStyle();
                        style.Number = 14; // Short date format
                        cell.SetStyle(style);
                    }
                }
            }

            // Save the workbook as an Excel file
            workbook.Save(excelPath, SaveFormat.Xlsx);
            Console.WriteLine($"Excel file saved to {excelPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
