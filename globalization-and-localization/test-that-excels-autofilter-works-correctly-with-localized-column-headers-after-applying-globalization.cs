// Title: Testing Aspose.Cells auto‑filter with French‑language column headers in a C# workbook
// Common Searches: how to test Aspose.Cells auto filter with French column headers in C# | filter rows by localized header after setting workbook culture to fr-FR using Aspose.Cells | verify auto‑filter visibility for accented column names in a .NET Excel file | apply globalization to an Aspose.Cells workbook and use AutoFilter on French headers | use Aspose.Cells AutoFilter.Filter method to select rows where Age equals 30 in French locale
// Tags: auto-filter operation on French-language columns Aspose.Cells | workbook CultureInfo globalization Aspose.Cells .NET | row visibility validation after applying AutoFilter Aspose.Cells | saving workbook to MemoryStream as XLSX with Aspose.Cells | using AutoFilter.Filter method for column value filtering Aspose.Cells C#

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, sets its CultureInfo to French (fr-FR), adds French headers (Nom, Âge, Ville) with sample data, applies an AutoFilter to range A1:C4, filters the Age column for the value "30" using AutoFilter.Filter, checks that only the matching rows remain visible, outputs the test result, and saves the workbook to a MemoryStream in XLSX format.
class AutoFilterLocalizationTest
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Apply globalization (set culture to French - France)
            workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

            Worksheet sheet = workbook.Worksheets[0];

            // Localized column headers (French)
            sheet.Cells["A1"].PutValue("Nom");    // Name
            sheet.Cells["B1"].PutValue("Âge");    // Age
            sheet.Cells["C1"].PutValue("Ville"); // City

            // Sample data rows
            sheet.Cells["A2"].PutValue("Alice");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["C2"].PutValue("Paris");

            sheet.Cells["A3"].PutValue("Bob");
            sheet.Cells["B3"].PutValue(25);
            sheet.Cells["C3"].PutValue("Lyon");

            sheet.Cells["A4"].PutValue("Charlie");
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["C4"].PutValue("Marseille");

            // Apply auto‑filter to the range that includes headers and data
            sheet.AutoFilter.Range = "A1:C4";

            // Filter on the 'Âge' column (index 1, zero‑based) to show only rows where Age = 30
            // Using the Filter method which is available in Aspose.Cells
            sheet.AutoFilter.Filter(1, "30");

            // Verify visibility: rows 2 and 4 (Excel rows) should be visible, row 3 hidden
            bool row2Visible = !sheet.Cells.Rows[1].IsHidden; // Excel row 2
            bool row3Visible = !sheet.Cells.Rows[2].IsHidden; // Excel row 3
            bool row4Visible = !sheet.Cells.Rows[3].IsHidden; // Excel row 4

            if (row2Visible && row4Visible && !row3Visible)
            {
                Console.WriteLine("Auto‑filter works correctly with localized headers.");
            }
            else
            {
                Console.WriteLine("Auto‑filter test failed.");
            }

            // Save the workbook to a memory stream (no physical file needed)
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, SaveFormat.Xlsx);
                // The stream can be used for further processing or inspection if required.
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
