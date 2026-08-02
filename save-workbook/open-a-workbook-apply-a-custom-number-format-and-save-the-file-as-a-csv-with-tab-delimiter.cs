using System;
using Aspose.Cells;

namespace AsposeCellsCustomNumberFormatToTabCsv
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Apply a custom number format to a range (e.g., cells B2:B5)
            Style customStyle = workbook.CreateStyle();
            // Example: display numbers with two decimal places and thousand separator
            customStyle.Custom = "#,##0.00";
            // Apply the style to the desired range
            sheet.Cells.CreateRange("B2", "B5").ApplyStyle(customStyle, new StyleFlag { NumberFormat = true });

            // Configure text save options for CSV with tab delimiter
            TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv);
            saveOptions.Separator = '\t'; // Tab delimiter

            // Save the workbook as a tab‑delimited CSV file
            workbook.Save("output.tsv", saveOptions);
        }
    }
}