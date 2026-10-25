// Title: Generate an Excel workbook with a PivotTable and display the PivotTable ribbon using Aspose.Cells for .NET (C#)
// Common Searches: how to show the PivotTable ribbon in an Excel file generated with Aspose.Cells C# | Aspose.Cells C# create pivot table and enable built‑in pivot UI ribbon | add a PivotTable to a new worksheet and keep Excel pivot tools visible using Aspose.Cells | C# code to create pivot table and preserve PivotTable ribbon in saved .xlsx | Aspose.Cells example for displaying pivot ribbon after saving workbook
// Tags: Aspose.Cells create pivot table programmatically | show pivot ribbon in saved workbook | add row and data fields to pivot table C# | save workbook with pivot UI enabled | populate sample data for pivot table Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample creates a new workbook, fills sample data, adds a PivotTable on a separate sheet, assigns the Category column as a row field and the Amount column as a data field, and saves the file as PivotTableWithRibbon.xlsx so that the PivotTable ribbon is available when the workbook is opened.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("Food");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Travel");
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["A4"].PutValue("Food");
            sheet.Cells["B4"].PutValue(150);
            sheet.Cells["A5"].PutValue("Travel");
            sheet.Cells["B5"].PutValue(200);

            // Define the source data range as a string (e.g., "A1:B5")
            string sourceData = "A1:B5";

            // Add a new worksheet to host the pivot table
            int pivotSheetIndex = workbook.Worksheets.Add();
            Worksheet pivotSheet = workbook.Worksheets[pivotSheetIndex];
            pivotSheet.Name = "PivotReport";

            // Create the pivot table at cell A1 on the new sheet
            int pivotTableIndex = pivotSheet.PivotTables.Add("PivotTable1", "A1", sourceData);
            PivotTable pivotTable = pivotSheet.PivotTables[pivotTableIndex];

            // Add "Category" as a row field (field index 0) and "Amount" as a data field (field index 1)
            // In newer Aspose.Cells versions RowFields.Add expects a PivotField object.
            // Retrieve the source fields from the pivot table and add them to the respective collections.
            PivotField categoryField = pivotTable.RowFields[0];   // source field at index 0 (Category)
            PivotField amountField = pivotTable.DataFields[0];   // source field at index 1 (Amount)

            pivotTable.RowFields.Add(categoryField);
            pivotTable.DataFields.Add(amountField);

            // Save the workbook
            workbook.Save("PivotTableWithRibbon.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
