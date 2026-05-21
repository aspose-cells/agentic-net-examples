using System;
using Aspose.Cells;

class AutoFilterLocalizationTest
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Apply custom globalization settings (optional, demonstrates usage)
        SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();
        globalization.SetListSeparator(';');               // Example: change list separator
        workbook.Settings.GlobalizationSettings = globalization;

        // Localized column headers (e.g., French)
        cells["A1"].PutValue("Produit");   // "Product"
        cells["B1"].PutValue("Ventes");    // "Sales"

        // Sample data with localized product names
        cells["A2"].PutValue("Pomme");     // Apple
        cells["B2"].PutValue(120);
        cells["A3"].PutValue("Banane");    // Banana
        cells["B3"].PutValue(80);
        cells["A4"].PutValue("Orange");
        cells["B4"].PutValue(150);
        cells["A5"].PutValue("Pomme");     // Apple again
        cells["B5"].PutValue(200);

        // Define the autofilter range that includes the header row and data rows
        worksheet.AutoFilter.Range = "A1:B5";

        // Apply a filter on the first column (Produit) for the localized value "Pomme"
        worksheet.AutoFilter.Filter(0, "Pomme");
        worksheet.AutoFilter.Refresh();

        // Verify which rows are hidden after filtering
        // Row indices are zero‑based; row 0 is the header.
        for (int row = 1; row <= 5; row++) // rows 1‑5 correspond to Excel rows 2‑6
        {
            bool isHidden = worksheet.Cells.Rows[row].IsHidden;
            Console.WriteLine($"Row {row + 1} hidden: {isHidden}");
        }

        // Save the workbook to verify the filter result
        workbook.Save("AutoFilterLocalizationTest.xlsx");
    }
}