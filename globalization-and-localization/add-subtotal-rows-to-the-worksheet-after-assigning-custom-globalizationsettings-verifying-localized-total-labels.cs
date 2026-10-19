// Title: Insert German‑localized subtotal rows into an Excel worksheet using Aspose.Cells for .NET
// Common Searches: how to insert subtotal rows with a German label using Aspose.Cells in C# | Aspose.Cells set workbook culture to de-DE and add subtotals | programmatically add grouped subtotals to Excel with Aspose.Cells .NET | verify localized subtotal label after inserting rows with Aspose.Cells | C# example for inserting rows and calculating group sums in Aspose.Cells
// Tags: insert subtotal rows Aspose.Cells | German culture localization Aspose.Cells | grouped data subtotal Excel .NET | custom subtotal label Excel C# | Aspose.Cells workbook culture settings

using System;
using System.Globalization;
using Aspose.Cells;

// The example creates a workbook, populates category and amount data, applies German (de-DE) culture settings, iterates through rows to compute group totals, inserts a new row after each category with the German label "Summe" and the calculated sum, validates the first localized label, and saves the file as SubtotalLocalized.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];

            // Populate sample data (Category, Amount)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("A");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("B");
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["A5"].PutValue("B");
            sheet.Cells["B5"].PutValue(40);

            // Apply German culture for formatting/localization
            workbook.Settings.CultureInfo = new CultureInfo("de-DE");

            // Manually add subtotal rows (German label "Summe")
            int dataStartRow = 1;               // Zero‑based index of first data row
            int totalDataRows = 5;              // Number of data rows (including header row)
            int lastRow = dataStartRow + totalDataRows - 1; // Index of the last data row

            int row = dataStartRow;
            while (row <= lastRow)
            {
                string category = sheet.Cells[row, 0].StringValue;
                double groupSum = 0;
                int groupStart = row;

                // Accumulate rows belonging to the same category
                while (row <= lastRow && sheet.Cells[row, 0].StringValue == category)
                {
                    groupSum += sheet.Cells[row, 1].DoubleValue;
                    row++;
                }

                // Insert a subtotal row after the current group
                sheet.Cells.InsertRows(row, 1);
                sheet.Cells[row, 0].PutValue("Summe");   // German subtotal label
                sheet.Cells[row, 1].PutValue(groupSum);
                // Adjust the last row index because we added a new row
                lastRow++;
                // Move to the next row after the inserted subtotal
                row++;
            }

            // Verify that the first subtotal label is correctly localized ("Summe")
            var subtotalLabelCell = sheet.Cells[3, 0]; // Expected position after first group
            string label = subtotalLabelCell.StringValue;
            string expectedLabel = "Summe";

            if (label == expectedLabel)
            {
                Console.WriteLine("Subtotal label correctly localized: " + label);
            }
            else
            {
                Console.WriteLine($"Subtotal label localization mismatch. Expected '{expectedLabel}', got '{label}'.");
            }

            // Save the workbook
            string outputPath = "SubtotalLocalized.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
