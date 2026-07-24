// Title: Export Workbook to CSV with European Number Formatting (comma decimal, dot thousands) using Aspose.Cells C#
// Description: Demonstrates how to create a workbook, apply a custom number style, set the NumberDecimalSeparator to ',' and NumberGroupSeparator to '.' for European locales, and save the sheet as a CSV file with the correct decimal and grouping symbols.
// Keywords: Aspose.Cells | C# | .NET | CSV export | European number format | decimal comma | thousands separator | NumberDecimalSeparator | NumberGroupSeparator | custom number format | SaveFormat.Csv | locale settings
// Common Searches: Aspose.Cells export CSV with comma decimal separator | Set NumberDecimalSeparator in Aspose.Cells C# | European locale CSV output Aspose.Cells | Custom number format CSV Aspose.Cells .NET | Change thousand separator when saving CSV with Aspose.Cells
// Developer Intent: Save an Excel workbook as a CSV file that follows European numeric conventions (comma for decimals, dot for thousands).
// Use Cases: Generate CSV reports for German, French, or other EU markets where numbers use a comma as the decimal separator. | Create financial statements for EU regulators that require dot‑separated thousands. | Batch convert multiple Excel files to CSV while preserving locale‑specific number formatting. | Provide data feeds for ERP systems that expect European numeric notation.
// AI Prompts: Write C# code with Aspose.Cells to export a workbook to CSV using a comma as the decimal separator and a dot as the thousands separator. | Show how to apply a custom number format and configure NumberDecimalSeparator and NumberGroupSeparator before saving as CSV in Aspose.Cells .NET. | Provide an example that sets European locale settings for CSV export with Aspose.Cells and validates the output file.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, apply a custom number style, set the NumberDecimalSeparator to ',' and NumberGroupSeparator to '.' for European locales, and save the sheet as a CSV file with the correct decimal and grouping symbols.
    public class ExportWorkbookToCsvWithEuropeanFormatting
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Add a numeric value that will be formatted
                cells["A1"].PutValue(12345.6789);

                // Define a custom number format (e.g., "#,##0.00")
                Style style = workbook.CreateStyle();
                style.Custom = "#,##0.00";

                // Apply the style to the cell
                cells["A1"].SetStyle(style);

                // Set European locale separators: comma for decimal, dot for group
                workbook.Settings.NumberDecimalSeparator = ',';
                workbook.Settings.NumberGroupSeparator = '.';

                // Define output file path
                string outputPath = "EuropeanFormattedOutput.csv";

                // Save the workbook as CSV; the formatting respects the locale settings
                workbook.Save(outputPath, SaveFormat.Csv);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
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
