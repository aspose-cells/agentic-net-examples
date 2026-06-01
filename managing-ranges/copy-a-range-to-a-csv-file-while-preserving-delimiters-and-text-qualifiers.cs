using System;
using System.IO;
using Aspose.Cells;

namespace ExportRangeToCsvApp
{
    class ExportRangeToCsv
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample data (including a value that contains the delimiter)
                cells["A1"].PutValue("Name");
                cells["B1"].PutValue("Age");
                cells["C1"].PutValue("City");
                cells["A2"].PutValue("John, Doe");   // comma inside the text to test quoting
                cells["B2"].PutValue(30);
                cells["C2"].PutValue("New York");
                cells["A3"].PutValue("Alice");
                cells["B3"].PutValue(25);
                cells["C3"].PutValue("London");
                // Add an empty row to demonstrate separator handling for blank rows
                cells["A4"].PutValue(string.Empty);
                cells["B4"].PutValue(string.Empty);
                cells["C4"].PutValue(string.Empty);

                // Define the range that should be exported (A1:C3)
                Aspose.Cells.Range exportRange = cells.CreateRange("A1:C3");

                // Build a CellArea that represents the same range
                CellArea area = new CellArea
                {
                    StartRow = exportRange.FirstRow,
                    EndRow = exportRange.FirstRow + exportRange.RowCount - 1,
                    StartColumn = exportRange.FirstColumn,
                    EndColumn = exportRange.FirstColumn + exportRange.ColumnCount - 1
                };

                // Configure TxtSaveOptions to preserve delimiters and apply text qualifiers only when needed
                TxtSaveOptions saveOptions = new TxtSaveOptions
                {
                    Separator = ',',                     // Use comma as delimiter
                    SeparatorString = ",",               // Explicit string version of the delimiter
                    QuoteType = TxtValueQuoteType.Normal, // Quote only when special characters are present
                    ExportArea = area,                   // Export only the defined range
                    KeepSeparatorsForBlankRow = true     // Keep delimiters for completely blank rows
                };

                string outputPath = "ExportedRange.csv";

                // Save the selected range as a CSV file using the configured options
                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"File saved successfully: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}