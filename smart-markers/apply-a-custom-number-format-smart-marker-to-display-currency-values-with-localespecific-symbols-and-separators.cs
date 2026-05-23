using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerCurrency
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook wb = new Workbook();

                // Set the workbook culture to German (Germany) – this will affect currency symbols and separators
                wb.Settings.CultureInfo = new CultureInfo("de-DE");

                // Create a custom number format for currency (Euro) using an invariant pattern
                // The pattern will be adapted to the workbook culture when displayed
                Style currencyStyle = wb.CreateStyle();
                currencyStyle.Custom = "_-\"€\"* #,##0.00_-;_-\"€\"* -#,##0.00_-;_-\"€\"* \"-\"??_-;_-@_-";

                // Apply only the number format part of the style
                StyleFlag flag = new StyleFlag();
                flag.NumberFormat = true;

                // Define a range where the smart markers will be placed (first column A, rows 1‑5)
                Aspose.Cells.Range range = wb.Worksheets[0].Cells.CreateRange(0, 0, 5, 1);
                range.ApplyStyle(currencyStyle, flag);

                // Insert smart markers that will be replaced with actual numeric values later
                for (int row = 0; row < 5; row++)
                {
                    wb.Worksheets[0].Cells[row, 0].PutValue("&[Amount]");
                }

                // Save the workbook
                string outputPath = "SmartMarkerCurrency.xlsx";
                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}