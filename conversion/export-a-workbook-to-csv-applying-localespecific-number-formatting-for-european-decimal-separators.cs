using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsCsvExport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Add sample numeric data
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Price");
            cells["A2"].PutValue("Laptop");
            cells["B2"].PutValue(12345.67);
            cells["A3"].PutValue("Phone");
            cells["B3"].PutValue(987.65);

            // Apply European number formatting (comma as decimal separator, dot as group separator)
            workbook.Settings.NumberDecimalSeparator = ',';   // decimal separator
            workbook.Settings.NumberGroupSeparator = '.';    // thousands separator

            // Create a style that uses a numeric format (e.g., "#,##0.00")
            Style style = workbook.CreateStyle();
            style.Custom = "#,##0.00";
            cells["B2"].SetStyle(style);
            cells["B3"].SetStyle(style);

            // Configure CSV save options
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                // Use semicolon as field delimiter (common in European CSV files)
                Separator = ';',
                // Export the displayed value (respecting the custom number format)
                FormatStrategy = CellValueFormatStrategy.DisplayStyle,
                // Export only the active sheet (default behavior)
                ExportAllSheets = false
            };

            // Save the workbook as CSV using the configured options
            workbook.Save("EuropeanFormattedOutput.csv", csvOptions);
        }
    }
}