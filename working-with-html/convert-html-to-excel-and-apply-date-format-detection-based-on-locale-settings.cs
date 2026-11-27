// Title: Convert HTML to Excel in C# with Aspose.Cells and apply locale‑specific date formatting
// Common Searches: C# Aspose.Cells convert HTML file to XLSX with culture‑aware date parsing | How to detect and convert date strings from HTML to Excel using a specific locale in Aspose.Cells | Load HTML into Aspose.Cells workbook and apply locale based date format in .NET | Aspose.Cells example for parsing dates from HTML using CultureInfo in C# | Convert HTML tables to Excel while preserving date formats for en‑GB and en‑US
// Tags: Aspose.Cells HTML to XLSX with locale date parsing | C# apply culture-specific number format to Excel dates | detect and convert string dates in workbook cells | custom cell style based on CultureInfo | LoadOptions for HTML import in Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

namespace HtmlToExcelWithLocaleDateDetection
{
    // This C# example loads an HTML file into an Aspose.Cells Workbook, optionally uses a supplied locale (CultureInfo) to parse string values as dates, replaces them with true DateTime values, applies a locale‑specific custom number format, and saves the result as an XLSX file.
    class Program
    {
        /// <param name="args">
        /// args[0] - path to the source HTML file (required)
        /// args[1] - path to the output Excel file (required)
        /// args[2] - locale identifier (e.g., "en-US", "en-GB") (optional, defaults to current culture)
        /// </param>
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: HtmlToExcelWithLocaleDateDetection <input.html> <output.xlsx> [locale]");
                return;
            }

            string htmlPath = args[0];
            string excelPath = args[1];
            string locale = args.Length >= 3 ? args[2] : CultureInfo.CurrentCulture.Name;

            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: HTML file not found at '{htmlPath}'.");
                return;
            }

            // Load the HTML into a Workbook using LoadOptions.
            Workbook workbook;
            try
            {
                var loadOptions = new LoadOptions(LoadFormat.Html);
                workbook = new Workbook(htmlPath, loadOptions);
                Console.WriteLine("HTML successfully loaded into workbook.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during HTML loading: {ex.Message}");
                return;
            }

            // Prepare culture info for date parsing.
            CultureInfo culture;
            try
            {
                culture = new CultureInfo(locale);
                Console.WriteLine($"Using locale '{culture.Name}' for date detection.");
            }
            catch (CultureNotFoundException)
            {
                Console.WriteLine($"Warning: Locale '{locale}' not recognized. Falling back to invariant culture.");
                culture = CultureInfo.InvariantCulture;
            }

            // Determine the appropriate Excel date format string for the locale.
            string excelDateFormat = GetExcelDateFormatForCulture(culture);
            Console.WriteLine($"Excel date format to apply: \"{excelDateFormat}\"");

            // Iterate through all worksheets and cells to detect date strings.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        if (cell.Type == CellValueType.IsString)
                        {
                            string text = cell.StringValue?.Trim();
                            if (string.IsNullOrEmpty(text))
                                continue;

                            // Try to parse the string as a date using the specified culture.
                            if (DateTime.TryParse(text, culture, DateTimeStyles.None, out DateTime parsedDate))
                            {
                                // Replace the string with a true date value.
                                cell.PutValue(parsedDate);

                                // Apply the locale‑specific number format.
                                Style style = cell.GetStyle();
                                style.Custom = excelDateFormat;
                                cell.SetStyle(style);
                            }
                        }
                    }
                }
            }

            // Ensure the output directory exists.
            try
            {
                string? outputDir = Path.GetDirectoryName(excelPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not create output directory. {ex.Message}");
            }

            // Save the workbook.
            try
            {
                workbook.Save(excelPath);
                Console.WriteLine($"Workbook saved to '{excelPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }

        private static string GetExcelDateFormatForCulture(CultureInfo culture)
        {
            // US English uses month/day/year.
            if (culture.Name.Equals("en-US", StringComparison.OrdinalIgnoreCase))
                return "mm/dd/yyyy";

            // British English and many Commonwealth locales use day/month/year.
            if (culture.Name.Equals("en-GB", StringComparison.OrdinalIgnoreCase) ||
                (culture.Name.StartsWith("en-", StringComparison.OrdinalIgnoreCase) && !culture.Name.Equals("en-US", StringComparison.OrdinalIgnoreCase)))
                return "dd/mm/yyyy";

            // German, French, Spanish, Italian, etc. also commonly use day.month.year.
            if (culture.Name.StartsWith("de-", StringComparison.OrdinalIgnoreCase) ||
                culture.Name.StartsWith("fr-", StringComparison.OrdinalIgnoreCase) ||
                culture.Name.StartsWith("es-", StringComparison.OrdinalIgnoreCase) ||
                culture.Name.StartsWith("it-", StringComparison.OrdinalIgnoreCase))
                return "dd.mm.yyyy";

            // Default to ISO format which works universally.
            return "yyyy-mm-dd";
        }
    }
}
