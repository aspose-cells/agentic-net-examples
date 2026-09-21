// Title: Test Aspose.Cells AutoFilter on French column headers after setting workbook culture to fr-FR (C#)
// AI Prompts: Generate C# code that configures a Workbook's CultureInfo to fr-FR, creates French headers (Nom, Âge, Pays), applies an AutoFilter to A1:C4, filters the "Pays" column for "France", and prints the rows that remain visible. | Show how to programmatically verify which rows are hidden or visible after applying an AutoFilter on a localized header using Aspose.Cells.
// Common Searches: Aspose.Cells apply auto filter on French header column | C# set workbook cultureinfo fr-FR for Excel filtering with Aspose.Cells | filter Excel rows by value in non‑English column using Aspose.Cells .NET | how to check visible rows after auto filter in Aspose.Cells C# | globalization example for AutoFilter in Aspose.Cells
// Tags: Aspose.Cells auto filter localized headers | set workbook cultureinfo fr-FR Aspose.Cells | filter rows by column value Aspose.Cells | verify hidden rows after auto filter Aspose.Cells | Excel auto filter non‑english headers .NET | globalization auto filter Aspose.Cells

using System;
using System.Globalization;
using Aspose.Cells;

// The sample creates a new workbook, sets its CultureInfo to French (fr-FR), writes French headers (Nom, Âge, Pays) and sample data, applies an AutoFilter to the range A1:C4, adds a filter on the "Pays" column to show only rows where the country equals "France", iterates through the rows to output those that are not hidden, and saves the workbook as AutoFilterLocalizationTest.xlsx.
class AutoFilterLocalizationTest
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set workbook culture to French (France) to simulate globalization
            workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

            // Populate headers in French and some sample data
            // Headers: "Nom" (Name), "Âge" (Age), "Pays" (Country)
            sheet.Cells["A1"].PutValue("Nom");
            sheet.Cells["B1"].PutValue("Âge");
            sheet.Cells["C1"].PutValue("Pays");

            // Row 2
            sheet.Cells["A2"].PutValue("Alice");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["C2"].PutValue("France");

            // Row 3
            sheet.Cells["A3"].PutValue("Bob");
            sheet.Cells["B3"].PutValue(25);
            sheet.Cells["C3"].PutValue("USA");

            // Row 4
            sheet.Cells["A4"].PutValue("Claire");
            sheet.Cells["B4"].PutValue(28);
            sheet.Cells["C4"].PutValue("France");

            // Apply AutoFilter to the range containing headers and data
            sheet.AutoFilter.Range = "A1:C4";

            // Filter the "Pays" column (index 2) to show only rows where Country = "France"
            sheet.AutoFilter.AddFilter(2, "France"); // column index is zero‑based
            sheet.AutoFilter.Refresh();

            // Verify that only rows with "France" are visible
            Console.WriteLine("Visible rows after applying filter on localized header \"Pays\":");
            for (int row = 1; row <= sheet.Cells.MaxDataRow; row++) // start from row 1 (zero‑based) which is the second row in Excel
            {
                // Rows hidden by the filter have the IsHidden flag set to true
                if (!sheet.Cells.Rows[row].IsHidden)
                {
                    string name = sheet.Cells[row, 0].StringValue;
                    int age = (int)sheet.Cells[row, 1].IntValue;
                    string country = sheet.Cells[row, 2].StringValue;
                    Console.WriteLine($"Row {row + 1}: {name}, {age}, {country}");
                }
            }

            // Save the workbook to a file (optional, demonstrates lifecycle usage)
            workbook.Save("AutoFilterLocalizationTest.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
