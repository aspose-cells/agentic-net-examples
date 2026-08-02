// Title: C# – Convert HTML to Excel with locale‑aware date detection using Aspose.Cells
// Description: Loads an HTML file into an Aspose.Cells Workbook, scans each used cell, parses string values as dates using the system CultureInfo, replaces successful parses with true DateTime cells, applies the locale's short‑date format, and saves the result as an XLSX file.
// Keywords: Aspose.Cells HTML to Excel | C# convert HTML to XLSX | locale date parsing .NET | CultureInfo date detection | short date format Aspose.Cells | date conversion workbook | Excel date style C# | Aspose.Cells SaveFormat.Xlsx | HTML table to Excel dates
// Common Searches: Aspose.Cells convert HTML to Excel with date detection | C# parse dates from HTML using CultureInfo in Aspose.Cells | apply locale short date format after HTML to XLSX conversion | detect and format dates when loading HTML into a workbook | how to set custom date format based on system culture in Aspose.Cells
// Developer Intent: Convert an HTML document to an Excel workbook while automatically turning date strings into proper DateTime cells formatted according to the current locale.
// Use Cases: Migrate web‑based reports that contain date strings into Excel for calculation and charting. | Generate locale‑specific Excel files from HTML templates so international users see dates in their familiar format. | Automate bulk conversion of HTML tables to XLSX while preserving date semantics for downstream data processing.
// AI Prompts: Show a C# Aspose.Cells example that loads an HTML file, detects dates with CultureInfo, applies the locale short‑date pattern, and saves as XLSX. | Explain how to customize the number format after conversion when the target culture uses a non‑standard date pattern. | Provide robust error handling for missing HTML files and strings that cannot be parsed as dates during the conversion.

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

namespace HtmlToExcelWithLocaleDateDetection
{
    // Loads an HTML file into an Aspose.Cells Workbook, scans each used cell, parses string values as dates using the system CultureInfo, replaces successful parses with true DateTime cells, applies the locale's short‑date format, and saves the result as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Paths for source HTML and destination Excel files
            string htmlFilePath = "input.html";
            string excelFilePath = "output.xlsx";

            // Verify that the input HTML file exists
            if (!File.Exists(htmlFilePath))
            {
                Console.WriteLine($"Error: The HTML file '{htmlFilePath}' was not found.");
                return;
            }

            try
            {
                // Load the HTML file into a workbook
                Workbook workbook = new Workbook(htmlFilePath);

                // Get the first worksheet (adjust if needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Use the current culture for date parsing
                CultureInfo locale = CultureInfo.CurrentCulture;

                // Iterate through all used cells
                Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
                int lastRow = usedRange.FirstRow + usedRange.RowCount - 1;
                int lastCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                for (int row = usedRange.FirstRow; row <= lastRow; row++)
                {
                    for (int col = usedRange.FirstColumn; col <= lastCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];

                        // Process only string cells
                        if (cell.Type == CellValueType.IsString)
                        {
                            string text = cell.StringValue.Trim();

                            // Try to parse the string as a date using the locale settings
                            if (DateTime.TryParse(text, locale, DateTimeStyles.None, out DateTime parsedDate))
                            {
                                // Replace the string with a true DateTime value
                                cell.PutValue(parsedDate);

                                // Apply a date format that matches the locale (short date pattern)
                                Style style = cell.GetStyle();
                                style.Number = 14; // Built‑in short date format
                                style.Custom = locale.DateTimeFormat.ShortDatePattern;
                                cell.SetStyle(style);
                            }
                        }
                    }
                }

                // Save the workbook as an Excel file
                workbook.Save(excelFilePath, SaveFormat.Xlsx);

                Console.WriteLine($"HTML file '{htmlFilePath}' has been converted to Excel with locale‑aware date detection.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
