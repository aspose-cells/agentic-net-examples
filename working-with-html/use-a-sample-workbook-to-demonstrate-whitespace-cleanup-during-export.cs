using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
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

                // Add data starting from C3, leaving leading blank rows and columns
                sheet.Cells["C3"].PutValue("Data1");
                sheet.Cells["D4"].PutValue("Data2");
                sheet.Cells["E5"].PutValue("Data3");

                // Export to CSV without trimming leading blanks
                TxtSaveOptions noTrimOptions = new TxtSaveOptions
                {
                    TrimLeadingBlankRowAndColumn = false // keep leading blanks
                };
                workbook.Save("output_without_trim.csv", noTrimOptions);

                // Export to CSV with trimming leading blanks
                TxtSaveOptions trimOptions = new TxtSaveOptions
                {
                    TrimLeadingBlankRowAndColumn = true // remove leading blanks
                };
                workbook.Save("output_with_trim.csv", trimOptions);

                // Remove unused styles to further reduce file size
                workbook.RemoveUnusedStyles();

                // Export to HTML while removing bogus bottom row data
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    ExportBogusRowData = false // exclude bogus row data
                };
                workbook.Save("output.html", htmlOptions);

                Console.WriteLine("Export completed. Compare the CSV files to see whitespace trimming.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            WhitespaceCleanupDemo.Run();
        }
    }
}