// Title: Set workbook CultureInfo to French before creating localized subtotal rows with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set workbook CultureInfo to French before adding subtotals | how to localize subtotal labels in Excel using Aspose.Cells | apply fr-FR globalization settings in Aspose.Cells workbook | example of manual subtotal insertion with culture-specific formatting in Aspose.Cells
// Tags: Aspose.Cells workbook locale configuration | localized subtotal rows Aspose.Cells | French number formatting Aspose.Cells | manual subtotal calculation C# Aspose.Cells | globalization settings Excel Aspose.Cells

using System;
using System.Collections.Generic;
using System.Globalization;
using Aspose.Cells;

// The example creates a new Workbook, assigns a French CultureInfo via GlobalizationSettings, manually computes and inserts subtotal rows grouped by a category column, and saves the Excel file with localized labels and number formatting.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("Food");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Food");
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["A4"].PutValue("Transport");
            sheet.Cells["B4"].PutValue(50);
            sheet.Cells["A5"].PutValue("Transport");
            sheet.Cells["B5"].PutValue(70);

            // Apply French culture for number formatting
            workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

            // Add subtotals manually (group by Category column A, sum Amount column B)
            AddSubtotals(sheet, 0, 1, 0, 1);

            // Save the workbook
            string outputPath = "SubtotalsLocalized.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    /// <param name="sheet">Target worksheet.</param>
    /// <param name="headerRow">Zero‑based index of the header row.</param>
    /// <param name="dataStartRow">Zero‑based index of the first data row.</param>
    /// <param name="groupColumn">Zero‑based index of the column to group by.</param>
    /// <param name="sumColumn">Zero‑based index of the column to sum.</param>
    private static void AddSubtotals(Worksheet sheet, int headerRow, int dataStartRow, int groupColumn, int sumColumn)
    {
        // Read all rows into a list for processing
        int maxRow = sheet.Cells.MaxDataRow;
        var rows = new List<(string Category, double Amount)>();

        for (int r = dataStartRow; r <= maxRow; r++)
        {
            string cat = sheet.Cells[r, groupColumn].StringValue;
            double amt = sheet.Cells[r, sumColumn].DoubleValue;
            rows.Add((cat, amt));
        }

        // Prepare to insert rows from bottom to top to keep indices valid
        string previousCategory = null;
        double groupSum = 0;
        int insertPos = maxRow + 1; // start after the last data row

        for (int i = rows.Count - 1; i >= 0; i--)
        {
            var (cat, amt) = rows[i];
            if (previousCategory == null)
            {
                previousCategory = cat;
                groupSum = amt;
                continue;
            }

            if (cat != previousCategory)
            {
                // Insert subtotal row for the previous group
                sheet.Cells.InsertRows(insertPos, 1);
                sheet.Cells[insertPos, groupColumn].PutValue($"{previousCategory} Subtotal");
                sheet.Cells[insertPos, sumColumn].PutValue(groupSum);
                insertPos++;

                // Reset for new group
                previousCategory = cat;
                groupSum = amt;
            }
            else
            {
                groupSum += amt;
            }
        }

        // Insert subtotal for the first group
        sheet.Cells.InsertRows(insertPos, 1);
        sheet.Cells[insertPos, groupColumn].PutValue($"{previousCategory} Subtotal");
        sheet.Cells[insertPos, sumColumn].PutValue(groupSum);
    }
}
