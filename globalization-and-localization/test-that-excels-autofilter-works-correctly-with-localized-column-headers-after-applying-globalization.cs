// Title: C# – Verify Aspose.Cells AutoFilter with French Headers after Globalization Settings
// Description: Creates a workbook with French column headers (Produit, Ventes), applies a custom list separator via SettableGlobalizationSettings, sets an AutoFilter on A1:B5, filters the first column for "Pomme", reports hidden rows, clears the filter, and saves the file as AutoFilterLocalizationTest.xlsx.
// Keywords: Aspose.Cells | C# AutoFilter localization | French column headers Excel | SettableGlobalizationSettings list separator | filter hidden rows Aspose | globalization Excel .NET | auto filter test code
// Common Searches: Aspose.Cells filter French headers C# | AutoFilter with localized column names .NET | Set list separator globalization Aspose.Cells | How to check hidden rows after AutoFilter | Remove AutoFilter programmatically Aspose
// Developer Intent: Confirm that AutoFilter respects localized (French) headers and custom globalization settings when filtering rows.
// Use Cases: Generate a worksheet with French headers and filter products by name. | Programmatically detect which rows are hidden after applying a filter to validate localization handling. | Reset the AutoFilter to display all rows after verification.
// AI Prompts: Write a unit test that asserts rows 2 and 4 are visible and rows 3 and 5 are hidden after filtering "Pomme" on the French header "Produit" using Aspose.Cells. | Provide a reusable method that returns a list of hidden row indices for any filtered worksheet with localized headers. | Explain how to configure SettableGlobalizationSettings to change the list separator for CSV export while preserving AutoFilter functionality.

using System;
using Aspose.Cells;

namespace AutoFilterLocalizationTest
{
    // Creates a workbook with French column headers (Produit, Ventes), applies a custom list separator via SettableGlobalizationSettings, sets an AutoFilter on A1:B5, filters the first column for "Pomme", reports hidden rows, clears the filter, and saves the file as AutoFilterLocalizationTest.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Set localized column headers (French)
            cells["A1"].PutValue("Produit");   // "Product"
            cells["B1"].PutValue("Ventes");    // "Sales"

            // Populate sample data
            cells["A2"].PutValue("Pomme");     // Apple
            cells["B2"].PutValue(120);
            cells["A3"].PutValue("Banane");    // Banana
            cells["B3"].PutValue(80);
            cells["A4"].PutValue("Pomme");     // Apple
            cells["B4"].PutValue(150);
            cells["A5"].PutValue("Orange");
            cells["B5"].PutValue(200);

            // Apply globalization settings if needed (example: change list separator)
            SettableGlobalizationSettings gSettings = new SettableGlobalizationSettings();
            gSettings.SetListSeparator(';'); // just an example, not required for filter
            workbook.Settings.GlobalizationSettings = gSettings;

            // Define the autofilter range (including header row)
            sheet.AutoFilter.Range = "A1:B5";

            // Filter the first column (fieldIndex 0) for the value "Pomme"
            sheet.AutoFilter.Filter(0, "Pomme");
            sheet.AutoFilter.Refresh();

            // Verify which rows are hidden after filtering
            Console.WriteLine("Rows hidden after applying filter on localized header:");
            for (int row = 1; row <= sheet.Cells.MaxDataRow; row++) // zero‑based rows
            {
                bool isHidden = sheet.Cells.Rows[row].IsHidden;
                Console.WriteLine($"Row {row + 1}: Hidden = {isHidden}");
            }

            // Remove the filter and show all rows again
            sheet.AutoFilter.ShowAll();
            sheet.AutoFilter.Refresh();

            // Save the workbook (output file)
            workbook.Save("AutoFilterLocalizationTest.xlsx");
        }
    }
}
