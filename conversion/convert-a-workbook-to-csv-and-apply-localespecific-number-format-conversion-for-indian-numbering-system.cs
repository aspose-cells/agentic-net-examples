using System;
using Aspose.Cells;

namespace AsposeCellsIndianCsvConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source workbook (can be any supported Excel format)
            string sourcePath = "input.xlsx";

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Set the workbook region to India to use Indian locale settings
            workbook.Settings.Region = CountryCode.India;

            // Optionally override separators (default for India is ',' and '.')
            workbook.Settings.NumberGroupSeparator = ',';
            workbook.Settings.NumberDecimalSeparator = '.';

            // Create a style with Indian number grouping pattern
            Style indianStyle = workbook.CreateStyle();
            // Custom format: first group of three digits, then groups of two digits
            indianStyle.Custom = "#,##,##0";

            // Apply the style to all used cells in the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            StyleFlag flag = new StyleFlag { NumberFormat = true };
            sheet.Cells.ApplyStyle(indianStyle, flag);

            // Save the workbook as CSV using the provided Save method (rule)
            string csvPath = "output.csv";
            workbook.Save(csvPath, SaveFormat.Csv);

            Console.WriteLine($"Workbook converted to CSV with Indian number formatting: {csvPath}");
        }
    }
}