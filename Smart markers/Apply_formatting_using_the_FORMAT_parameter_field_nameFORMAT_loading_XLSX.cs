using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsFormattingExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Create LoadOptions specifying the XLSX format explicitly
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Access the first worksheet (you can change the index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the last row with data
            int maxRow = sheet.Cells.MaxDataRow;
            if (maxRow < 0) maxRow = 0; // ensure at least one row

            // ------------------------------------------------------------
            // Apply custom formatting to a column using the {{field_name:FORMAT}} concept.
            // Example: format column B as a date in "dd-MMM-yyyy" format.
            // ------------------------------------------------------------

            // Create a new style object for dates
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "dd-MMM-yyyy";

            // Apply the style to the entire column B (index 1)
            AsposeRange dateRange = sheet.Cells.CreateRange(0, 1, maxRow + 1, 1);
            dateRange.ApplyStyle(dateStyle, new StyleFlag { NumberFormat = true });

            // ------------------------------------------------------------
            // Another example: format column C as currency with two decimal places.
            // ------------------------------------------------------------
            Style currencyStyle = workbook.CreateStyle();
            currencyStyle.Custom = "\"$\"#,##0.00";

            // Apply the style to the entire column C (index 2)
            AsposeRange currencyRange = sheet.Cells.CreateRange(0, 2, maxRow + 1, 1);
            currencyRange.ApplyStyle(currencyStyle, new StyleFlag { NumberFormat = true });

            // Save the modified workbook to a new file
            string outputPath = "output_formatted.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook loaded from '{sourcePath}', formatted, and saved as '{outputPath}'.");
        }
    }
}