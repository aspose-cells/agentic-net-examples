// Title: Add a calculated field to an Aspose.Cells pivot table and refresh results with PivotTable.CalculateData in C#
// AI Prompts: Create a pivot table from a range, define a custom calculated field that multiplies two columns, then call RefreshData and CalculateData using Aspose.Cells for .NET. | Update an existing Aspose.Cells pivot table after inserting a new calculated field by invoking the CalculateData method in C#. | Generate an Excel workbook that contains a pivot table with a "Total" calculated field (Quantity × Price) and ensure the values are computed automatically.
// Common Searches: Aspose.Cells add calculated field to pivot table C# example | how to use PivotTable.CalculateData after adding a calculated field in Aspose.Cells | refresh pivot cache and recalculate data with Aspose.Cells .NET | create Excel pivot with custom formula using Aspose.Cells C#
// Tags: add calculated field to Aspose.Cells pivot table | PivotTable.CalculateData method C# | refresh pivot cache Aspose.Cells | custom formula in Excel pivot Aspose.Cells | populate pivot data after calculated field Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// // This program creates a workbook, fills it with product data, adds a pivot table, defines a calculated field "Total" as Quantity*Price, refreshes the pivot cache, calculates the data, and saves the workbook as PivotWithCalculatedField.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate source data for the pivot table
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Quantity");
        sheet.Cells["C1"].PutValue("Price");

        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["C2"].PutValue(2);

        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue(5);
        sheet.Cells["C3"].PutValue(3);

        sheet.Cells["A4"].PutValue("Apple");
        sheet.Cells["B4"].PutValue(7);
        sheet.Cells["C4"].PutValue(2);

        // Add a pivot table based on the data range A1:C4, place it at E3, and name it "SalesPivot"
        int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "SalesPivot");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Configure the pivot table: Product as row field, Quantity and Price as data fields
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Price");

        // Add a calculated field "Total" that multiplies Quantity by Price and drag it to the data area
        pivotTable.AddCalculatedField("Total", "=Quantity*Price", true);

        // Refresh the pivot cache and calculate the data so that the calculated field values are populated
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook with the pivot table and calculated field
        workbook.Save("PivotWithCalculatedField.xlsx");
    }
}
