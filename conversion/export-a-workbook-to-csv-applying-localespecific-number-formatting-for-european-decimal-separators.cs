// Title: Export a C# Aspose.Cells workbook to CSV using European number formatting (comma decimal separator and semicolon delimiter)
// AI Prompts: Generate C# code that creates a workbook, sets NumberDecimalSeparator to ',' and NumberGroupSeparator to '.', applies a custom numeric style "#,##0.00", and saves the file as CSV with ';' separator using TxtSaveOptions. | Show how to configure Aspose.Cells TxtSaveOptions to preserve the cell's display style when exporting to CSV with a European locale. | Write a C# snippet that applies a numeric style to cells and exports the workbook to a semicolon‑delimited CSV file with comma decimal separators.
// Common Searches: Aspose.Cells C# export to CSV with semicolon delimiter and comma decimal separator | How to set European number format when saving workbook as CSV using Aspose.Cells | Save workbook as CSV preserving cell display style Aspose.Cells .NET | Configure TxtSaveOptions for CSV with custom number formatting in Aspose.Cells | C# Aspose.Cells CSV output using European locale settings
// Tags: Aspose.Cells CSV export with European locale | TxtSaveOptions semicolon delimiter | NumberDecimalSeparator comma Aspose.Cells | DisplayStyle format strategy CSV | custom numeric format #,##0.00 Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsExamples
{
    // The example creates a workbook, configures the number decimal separator to a comma and the group separator to a dot, defines a custom numeric style "#,##0.00", populates sample cells, and saves the workbook as a CSV file using a semicolon field separator while preserving the display style for European number formatting.
    public class ExportWorkbookToCsvWithEuropeanFormatting
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Set European locale number formatting (comma as decimal separator, dot as group separator)
                workbook.Settings.NumberDecimalSeparator = ',';
                workbook.Settings.NumberGroupSeparator = '.';

                // Create a style that uses a numeric format with two decimal places
                Style europeanStyle = workbook.CreateStyle();
                europeanStyle.Custom = "#,##0.00";

                // Populate some sample numeric data and apply the style
                Worksheet sheet = workbook.Worksheets[0];
                Cell cellA1 = sheet.Cells["A1"];
                cellA1.PutValue(12345.6789);
                cellA1.SetStyle(europeanStyle);

                Cell cellA2 = sheet.Cells["A2"];
                cellA2.PutValue(9876.5432);
                cellA2.SetStyle(europeanStyle);

                // Configure CSV save options:
                // - Use semicolon as field separator (common in Europe when decimal separator is a comma)
                // - Apply the cell's display style when converting values to text
                TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
                {
                    Separator = ';',
                    FormatStrategy = CellValueFormatStrategy.DisplayStyle
                };

                // Save the workbook as a CSV file using the configured options
                string outputPath = "EuropeanNumbers.csv";
                workbook.Save(outputPath, csvOptions);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
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
