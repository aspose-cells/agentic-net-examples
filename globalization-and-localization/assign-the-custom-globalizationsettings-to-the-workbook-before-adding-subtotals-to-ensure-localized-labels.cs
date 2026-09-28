// Title: Set French GlobalizationSettings on an Aspose.Cells workbook before generating category subtotals in C#
// AI Prompts: Apply a French CultureInfo to the workbook's Settings.CultureInfo, then calculate and insert subtotal rows for each category using Aspose.Cells. | Create a C# routine that groups rows by a column, sums numeric values, and writes localized subtotal rows with number formatting based on GlobalizationSettings.
// Common Searches: aspnet how to set workbook culture to fr-FR before adding subtotal rows with Aspose.Cells | c# Aspose.Cells localized number formatting for subtotal calculations | example of grouping data and inserting subtotals after setting GlobalizationSettings in Aspose.Cells
// Tags: Aspose.Cells set workbook culture info | C# generate localized subtotals in Excel | globalizationsettings fr-FR Aspose.Cells | group by column and sum using Aspose.Cells API | Excel export with French number formatting Aspose.Cells

using System;
using System.Collections.Generic;
using System.Globalization;
using Aspose.Cells;

// The example creates a workbook, fills it with category and amount data, sets the workbook's CultureInfo to French (fr-FR) to localize number formatting, computes totals per category, writes subtotal rows, and saves the file as SubtotalsWithLocalization.xlsx using Aspose.Cells for .NET.
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
            sheet.Cells["B2"].PutValue(120.5);
            sheet.Cells["A3"].PutValue("Food");
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["A4"].PutValue("Transport");
            sheet.Cells["B4"].PutValue(50);
            sheet.Cells["A5"].PutValue("Transport");
            sheet.Cells["B5"].PutValue(70);

            // Apply French culture settings for number formatting
            workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

            // Calculate subtotals manually (group by Category in column A, sum Amount in column B)
            int firstDataRow = 1; // zero‑based index, row after header
            int lastDataRow = sheet.Cells.MaxDataRow;
            var sums = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);

            for (int row = firstDataRow; row <= lastDataRow; row++)
            {
                string category = sheet.Cells[row, 0].StringValue;
                double amount = sheet.Cells[row, 1].DoubleValue;
                if (!sums.ContainsKey(category))
                    sums[category] = 0;
                sums[category] += amount;
            }

            // Write subtotal rows after the original data
            int subtotalStartRow = lastDataRow + 2; // leave one empty row
            int currentRow = subtotalStartRow;
            foreach (var kvp in sums)
            {
                sheet.Cells[currentRow, 0].PutValue($"{kvp.Key} Total");
                sheet.Cells[currentRow, 1].PutValue(kvp.Value);
                currentRow++;
            }

            // Save the workbook
            string outputPath = "SubtotalsWithLocalization.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
