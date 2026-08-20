// Title: C# – Control Whitespace When Exporting to CSV with Aspose.Cells TxtSaveOptions
// Description: Creates a workbook, places data at C3, and saves it twice as CSV—once preserving leading empty rows/columns and once removing them—by toggling TxtSaveOptions.TrimLeadingBlankRowAndColumn. The example lets you compare the two outputs to see the whitespace‑cleanup effect.
// Keywords: Aspose.Cells | C# | TxtSaveOptions | TrimLeadingBlankRowAndColumn | CSV export | whitespace cleanup | preserve leading blanks | remove empty rows | Excel to CSV | Aspose.Cells example
// Common Searches: How to keep empty rows when saving CSV with Aspose.Cells | Aspose.Cells TxtSaveOptions TrimLeadingBlankRowAndColumn false example | Export Excel to CSV without trimming blanks C# | Difference between TrimLeadingBlankRowAndColumn true and false | Aspose.Cells whitespace removal CSV | C# code to compare CSV outputs with and without leading blanks
// Developer Intent: Show how to enable or disable whitespace trimming during CSV export by setting the TrimLeadingBlankRowAndColumn property of TxtSaveOptions.
// Use Cases: Generate CSV reports that retain the original worksheet layout for tools that rely on fixed row/column positions. | Produce compact CSV files for data‑import pipelines that require dense, blank‑free content. | Validate the impact of whitespace trimming by comparing two CSV files created with opposite TrimLeadingBlankRowAndColumn settings.
// AI Prompts: Write C# code using Aspose.Cells to export a workbook to CSV while preserving leading blank rows and columns. | Explain how TxtSaveOptions.TrimLeadingBlankRowAndColumn influences CSV output and how to switch it on or off. | Provide step‑by‑step instructions to compare CSV files generated with and without whitespace trimming.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, places data at C3, and saves it twice as CSV—once preserving leading empty rows/columns and once removing them—by toggling TxtSaveOptions.TrimLeadingBlankRowAndColumn. The example lets you compare the two outputs to see the whitespace‑cleanup effect.
    public class WhitespaceCleanupDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add data starting from cell C3 (leaving leading blank rows and columns)
                sheet.Cells["C3"].PutValue("Data1");
                sheet.Cells["D4"].PutValue("Data2");
                sheet.Cells["E5"].PutValue("Data3");

                // Export without trimming leading blanks (default is true, set to false explicitly)
                TxtSaveOptions optionsNoTrim = new TxtSaveOptions
                {
                    TrimLeadingBlankRowAndColumn = false // keep leading blanks
                };
                workbook.Save("output_without_trim.csv", optionsNoTrim);

                // Export with trimming leading blanks (default behavior)
                TxtSaveOptions optionsTrim = new TxtSaveOptions
                {
                    TrimLeadingBlankRowAndColumn = true // remove leading blanks
                };
                workbook.Save("output_with_trim.csv", optionsTrim);

                Console.WriteLine("Export completed. Compare 'output_without_trim.csv' and 'output_with_trim.csv' to see whitespace cleanup.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during processing: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                WhitespaceCleanupDemo.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
