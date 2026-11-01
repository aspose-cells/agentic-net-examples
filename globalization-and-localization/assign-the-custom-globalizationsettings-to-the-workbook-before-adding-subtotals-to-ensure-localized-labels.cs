// Title: How to apply German CultureInfo to an Aspose.Cells workbook before inserting subtotal rows for localized labels (C#)
// Common Searches: Aspose.Cells set workbook cultureinfo to de-DE for German subtotal labels in C# | C# add subtotal rows with localized labels using Aspose.Cells workbook settings | How to localize Excel subtotal text to German with Aspose.Cells .NET | Insert subtotal rows after grouping data with culture-specific labels in Aspose.Cells | Saving workbook with German UI strings for subtotal rows using Aspose.Cells
// Tags: workbook cultureinfo Aspose.Cells C# | localized subtotal rows Excel Aspose.Cells | German subtotal label Summe Aspose.Cells | grouped data subtotal insertion Aspose.Cells | globalization settings workbook Aspose.Cells

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, assigns the German CultureInfo (de-DE) to the workbook settings, populates sample category and amount data, groups rows by category, inserts subtotal rows with German labels such as "Kategorie Summe" and SUM formulas, and saves the file as LocalizedSubtotals.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set German culture for localized subtotal labels (e.g., "Summe")
            workbook.Settings.CultureInfo = new CultureInfo("de-DE");

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data
            cells["A1"].PutValue("Kategorie"); // Category header (German)
            cells["B1"].PutValue("Betrag");    // Amount header (German)

            cells["A2"].PutValue("Essen");
            cells["B2"].PutValue(12.5);
            cells["A3"].PutValue("Essen");
            cells["B3"].PutValue(8.75);
            cells["A4"].PutValue("Transport");
            cells["B4"].PutValue(15.0);
            cells["A5"].PutValue("Transport");
            cells["B5"].PutValue(20.0);

            // Manually add subtotal rows (sum) for each category
            // Collect rows per category
            var categoryRows = new Dictionary<string, List<int>>();
            for (int row = 1; row <= 4; row++) // data rows are 2..5 (zero‑based index)
            {
                string category = cells[row, 0].StringValue;
                if (!categoryRows.ContainsKey(category))
                    categoryRows[category] = new List<int>();
                categoryRows[category].Add(row);
            }

            // Insert subtotal rows starting from the bottom to keep indices valid
            int offset = 0;
            foreach (var kvp in categoryRows)
            {
                // Determine insertion point: after the last row of the current group
                int insertRow = kvp.Value[kvp.Value.Count - 1] + 1 + offset;

                // Insert a new row for the subtotal
                sheet.Cells.InsertRows(insertRow, 1);
                offset++; // account for the newly inserted row

                // Write label and sum formula
                cells[insertRow, 0].PutValue($"{kvp.Key} Summe");
                // Sum the amount column for the group's rows
                int firstDataRow = kvp.Value[0] + 1; // Excel rows are 1‑based in formulas
                int lastDataRow = kvp.Value[kvp.Value.Count - 1] + 1;
                string sumFormula = $"SUM(B{firstDataRow}:B{lastDataRow})";
                cells[insertRow, 1].Formula = sumFormula;
            }

            // Determine output file path
            string outputPath = "LocalizedSubtotals.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
