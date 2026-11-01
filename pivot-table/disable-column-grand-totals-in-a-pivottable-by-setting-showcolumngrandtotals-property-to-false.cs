// Title: Hide column grand totals in an Aspose.Cells PivotTable using C#
// Common Searches: Aspose.Cells C# hide column grand totals in pivot table | disable column grand totals for Excel pivot table with Aspose.Cells .NET | set PivotTable.ColumnGrand false in C# Aspose.Cells example | remove column grand total from generated Excel pivot using Aspose.Cells | programmatically hide column grand totals in Excel pivot via Aspose.Cells library
// Tags: Aspose.Cells PivotTable ColumnGrand property | disable column grand totals Aspose.Cells | C# hide Excel pivot column totals | Excel pivot table formatting Aspose.Cells | programmatic pivot table grand totals C#

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;

// The example creates a workbook, adds sample data, builds a PivotTable on a new sheet, assigns the Category field to rows and the Amount field to data, disables column grand totals while keeping row grand totals visible, and saves the workbook as an .xlsx file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and populate sample data
            Workbook wb = new Workbook();
            Worksheet dataSheet = wb.Worksheets[0];
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Amount");
            dataSheet.Cells["A2"].PutValue("A");
            dataSheet.Cells["B2"].PutValue(100);
            dataSheet.Cells["A3"].PutValue("B");
            dataSheet.Cells["B3"].PutValue(200);
            dataSheet.Cells["A4"].PutValue("A");
            dataSheet.Cells["B4"].PutValue(150);
            dataSheet.Cells["A5"].PutValue("B");
            dataSheet.Cells["B5"].PutValue(250);

            // Add a worksheet to host the pivot table
            int pivotSheetIndex = wb.Worksheets.Add();
            Worksheet pivotSheet = wb.Worksheets[pivotSheetIndex];

            // Define source data range
            string sourceData = "=Sheet1!A1:B5";

            // Add the pivot table and obtain its index
            int pivotTableIndex = pivotSheet.PivotTables.Add(sourceData, "A1", "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotTableIndex];

            // Add "Category" to rows (field index 0) and "Amount" to data area (field index 1)
            pivotTable.RowFields.Add(pivotTable.RowFields[0]);   // Category
            pivotTable.DataFields.Add(pivotTable.DataFields[0]); // Amount

            // Disable column grand totals, keep row grand totals enabled
            pivotTable.ColumnGrand = false; // hide column grand totals
            pivotTable.RowGrand = true;     // show row grand totals

            // Save the workbook
            string outputPath = "PivotTable_NoColumnGrandTotals.xlsx";
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
