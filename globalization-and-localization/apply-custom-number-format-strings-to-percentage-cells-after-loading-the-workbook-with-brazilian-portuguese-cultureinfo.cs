using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsPercentageFormatting
{
    class Program
    {
        static void Main()
        {
            // Load the workbook with Brazilian Portuguese culture (pt-BR)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            loadOptions.CultureInfo = new CultureInfo("pt-BR");
            Workbook workbook = new Workbook("input.xlsx", loadOptions);

            // Iterate through all cells in the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            foreach (Cell cell in cells)
            {
                // Retrieve the cell's style
                Style style = cell.GetStyle();

                // Check if the current style is a percentage format
                if (style.IsPercent)
                {
                    // Apply a custom percentage format that respects the culture
                    // "#,##0.00%" will use the culture's decimal and group separators
                    style.SetCustom("#,##0.00%", true);

                    // Re-apply the modified style to the cell
                    cell.SetStyle(style);
                }
            }

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}