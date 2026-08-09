// Title: Export Workbook to CSV with European Number Formatting and Semicolon Delimiter (C# Aspose.Cells)
// Description: Creates a workbook, sets ',' as decimal and '.' as thousands separator, adds sample data, configures TxtSaveOptions with a semicolon separator and DisplayStyle format strategy, then saves the file so numeric values appear in European style (e.g., 12.345,67).
// Keywords: Aspose.Cells CSV export | European number format | comma decimal separator | semicolon delimiter | TxtSaveOptions | DisplayStyle strategy | C# .NET | locale specific CSV | NumberGroupSeparator | SaveFormat.Csv
// Common Searches: Aspose.Cells export CSV European format C# | How to set comma as decimal separator in Aspose.Cells | Semicolon delimited CSV with Aspose.Cells | Save Excel as CSV with custom separators .NET | DisplayStyle format strategy for CSV export
// Developer Intent: Save an Excel workbook as a CSV file that follows European conventions—comma decimal, dot thousands separator—and uses a semicolon as the field delimiter.
// Use Cases: Generating CSV reports for European markets where numbers must use a comma for decimals. | Creating data files for legacy ERP systems that expect semicolon‑delimited CSV with European numeric formatting. | Automating localized CSV exports from ASP.NET applications without manual string processing.
// AI Prompts: Show how to change the locale to French (comma decimal, space thousands) while exporting CSV with Aspose.Cells. | Explain how to export each worksheet to a separate CSV file while preserving European number formatting. | Provide code to include date and currency formatting together with the European numeric settings in a CSV export.

using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsExamples
{
    // Creates a workbook, sets ',' as decimal and '.' as thousands separator, adds sample data, configures TxtSaveOptions with a semicolon separator and DisplayStyle format strategy, then saves the file so numeric values appear in European style (e.g., 12.345,67).
    public class ExportWorkbookToCsvWithEuropeanFormatting
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Set European number formatting: comma as decimal separator, dot as group separator
                workbook.Settings.NumberDecimalSeparator = ',';
                workbook.Settings.NumberGroupSeparator = '.';

                // Add sample numeric data
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Price");
                sheet.Cells["A2"].PutValue("Laptop");
                sheet.Cells["B2"].PutValue(12345.67); // Will be formatted as 12.345,67
                sheet.Cells["A3"].PutValue("Phone");
                sheet.Cells["B3"].PutValue(987.65);   // Will be formatted as 987,65

                // Create CSV save options
                TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
                {
                    // Use semicolon as delimiter (common in European CSV files)
                    Separator = ';',
                    // Export numbers using the display style (applies the workbook's decimal/group separators)
                    FormatStrategy = CellValueFormatStrategy.DisplayStyle
                };

                // Save the workbook as CSV with the specified options
                string outputPath = "EuropeanFormattedOutput.csv";
                workbook.Save(outputPath, csvOptions);
                Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during export: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportWorkbookToCsvWithEuropeanFormatting.Run();
        }
    }
}
