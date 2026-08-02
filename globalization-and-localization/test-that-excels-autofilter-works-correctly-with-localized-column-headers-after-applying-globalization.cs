using System;
using Aspose.Cells;

namespace AutoFilterLocalizationTest
{
    // Custom globalization settings (optional, demonstrates applying globalization)
    public class CustomGlobalizationSettings : SettableGlobalizationSettings
    {
        public CustomGlobalizationSettings()
        {
            // Example: change the list separator to semicolon
            this.SetListSeparator(';');
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Apply custom globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

            // Populate localized column headers (e.g., French)
            cells["A1"].PutValue("Produit");   // "Product"
            cells["B1"].PutValue("Ventes");    // "Sales"

            // Add sample data rows
            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue(1200);
            cells["A3"].PutValue("Banane");
            cells["B3"].PutValue(850);
            cells["A4"].PutValue("Cerise");
            cells["B4"].PutValue(430);
            cells["A5"].PutValue("Datte");
            cells["B5"].PutValue(670);

            // Set the auto‑filter range (header row + data rows)
            // Parameters: startRow, startColumn, endRow
            // Here we filter column A (index 0) from row 0 to row 4 (5 rows total)
            sheet.AutoFilter.SetRange(0, 0, 4);

            // Apply a filter on the first column (fieldIndex 0) for the value "Banane"
            sheet.AutoFilter.Filter(0, "Banane");

            // Refresh the filter to hide rows that do not match the criteria
            sheet.AutoFilter.Refresh();

            // Verify that only the row with "Banane" remains visible
            // (Rows are zero‑based; row 2 contains "Banane")
            for (int row = 0; row <= sheet.Cells.MaxDataRow; row++)
            {
                bool hidden = sheet.Cells.Rows[row].IsHidden;
                Console.WriteLine($"Row {row + 1} hidden: {hidden}");
            }

            // Save the workbook to verify the result
            workbook.Save("AutoFilterLocalizedHeaders.xlsx");
        }
    }
}