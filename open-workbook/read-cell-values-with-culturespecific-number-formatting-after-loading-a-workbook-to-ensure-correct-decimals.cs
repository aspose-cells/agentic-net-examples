using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsCultureDemo
{
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a sample workbook and write a numeric value (dot as decimal)
            // -----------------------------------------------------------------
            Workbook sampleWorkbook = new Workbook();
            Worksheet sheet = sampleWorkbook.Worksheets[0];
            // Put a double value using the invariant dot separator
            sheet.Cells["A1"].PutValue(1234.56);
            // Save the workbook in XLSX format
            string samplePath = "sample.xlsx";
            sampleWorkbook.Save(samplePath, SaveFormat.Xlsx);

            // -----------------------------------------------------------------
            // 2. Load the workbook with a specific culture (German - uses comma)
            // -----------------------------------------------------------------
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            // Set the culture that defines the decimal separator
            loadOptions.CultureInfo = new CultureInfo("de-DE");

            // Load the workbook using the culture‑specific options
            Workbook workbook = new Workbook(samplePath, loadOptions);

            // -----------------------------------------------------------------
            // 3. Read the cell value using different properties
            // -----------------------------------------------------------------
            Cell cell = workbook.Worksheets[0].Cells["A1"];

            // DoubleValue gives the raw numeric value (independent of culture)
            double rawNumber = cell.DoubleValue;

            // StringValue returns the formatted string according to the loaded culture
            string formattedString = cell.StringValue;

            // Display the results
            Console.WriteLine($"Raw numeric value (DoubleValue): {rawNumber}");
            Console.WriteLine($"Formatted string value (StringValue) with German culture: {formattedString}");

            // -----------------------------------------------------------------
            // 4. Optionally, change workbook settings to use a custom decimal separator
            // -----------------------------------------------------------------
            workbook.Settings.NumberDecimalSeparator = ',';
            workbook.Settings.NumberGroupSeparator = '.';
            // Re‑format the cell using a custom style to see the effect
            Style style = workbook.CreateStyle();
            style.Custom = "#,##0.00";
            cell.SetStyle(style);

            // After applying the style, StringValue reflects the new separators
            Console.WriteLine($"After applying custom separators: {cell.StringValue}");
        }
    }
}