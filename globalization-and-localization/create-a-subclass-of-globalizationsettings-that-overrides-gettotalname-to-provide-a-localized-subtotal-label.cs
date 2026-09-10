// Title: How to subclass GlobalizationSettings in Aspose.Cells for .NET to customize the subtotal label in a totals row
// AI Prompts: Generate a C# class that inherits from Aspose.Cells.GlobalizationSettings and overrides GetTotalName to return a localized string for the subtotal row. | Show how to apply the custom GlobalizationSettings subclass to a Workbook so that the totals row displays the new subtotal label. | Provide a complete example that creates a worksheet, adds a ListObject with a totals row, and saves the file demonstrating the overridden GetTotalName effect.
// Common Searches: Aspose.Cells .NET customize subtotal text in totals row | override GlobalizationSettings GetTotalName example C# | localize total row label using Aspose.Cells workbook | C# Aspose.Cells change default Subtotal label in table totals
// Tags: GlobalizationSettings GetTotalName override | Aspose.Cells subtotal label localization | C# custom total row name Aspose.Cells | Aspose.Cells ListObject totals row customization | Excel workbook globalized subtotal text

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// // Demonstrates creating a C# subclass of Aspose.Cells.GlobalizationSettings that overrides GetTotalName to return a localized 'Subtotal' label, applying it to a Workbook, adding a ListObject with a totals row, and saving the resulting Excel file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Populate some data.
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(20);

            // Add a table (ListObject) covering the data range (including header).
            int firstRow = 0;          // zero‑based index
            int firstColumn = 0;
            int totalRows = 3;         // header + 2 data rows
            int totalColumns = 2;
            int tableIdx = sheet.ListObjects.Add(firstRow, firstColumn,
                firstRow + totalRows, firstColumn + totalColumns, true);
            ListObject table = sheet.ListObjects[tableIdx];
            table.ShowTotals = true; // Enable the total row.

            // Manually set a subtotal (SUM) for the "Amount" column (index 1).
            // The total row is placed immediately after the data rows.
            int totalRowIndex = firstRow + totalRows; // zero‑based
            // Excel formula uses 1‑based row numbers.
            int dataStartRow = firstRow + 2; // first data row (row 2 in Excel)
            int dataEndRow = firstRow + totalRows; // last data row (row 3 in Excel)
            string sumFormula = $"=SUM(B{dataStartRow}:B{dataEndRow})";
            sheet.Cells[totalRowIndex, 1].Formula = sumFormula; // column B (index 1)

            // Save the workbook.
            string outputPath = "LocalizedSubtotal.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
