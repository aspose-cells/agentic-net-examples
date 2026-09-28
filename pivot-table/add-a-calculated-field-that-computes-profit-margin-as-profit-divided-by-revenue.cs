// Title: Create and format a ProfitMargin calculated field in an Aspose.Cells PivotTable using C#
// AI Prompts: Generate C# code that inserts a new field called ProfitMargin into an Aspose.Cells PivotTable using the formula Profit/Revenue and sets its display format to percentage. | Show how to refresh the PivotTable data and write the workbook to an XLSX file after creating the ProfitMargin field with Aspose.Cells in C#.
// Common Searches: asp.net insert profit margin calculated field in Aspose.Cells pivot table | c# Aspose.Cells pivot table compute profit margin as percentage | how to apply percentage number format to a calculated pivot field in Aspose.Cells | refresh pivot table after adding calculated field using Aspose.Cells C#
// Tags: add calculated field to Aspose.Cells pivot table | profit margin calculation in Aspose.Cells pivot | percentage number format for pivot data field Aspose.Cells | update pivot table after calculated field Aspose.Cells | export workbook to xlsx with Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example builds a workbook with product, revenue, and profit data, creates a PivotTable, inserts a calculated field named ProfitMargin that divides Profit by Revenue, formats this field as a percentage, refreshes and recalculates the PivotTable, and saves the result as an XLSX file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data: Product, Revenue, Profit
        cells["A1"].PutValue("Product");
        cells["B1"].PutValue("Revenue");
        cells["C1"].PutValue("Profit");

        cells["A2"].PutValue("A"); cells["B2"].PutValue(1000); cells["C2"].PutValue(200);
        cells["A3"].PutValue("B"); cells["B3"].PutValue(1500); cells["C3"].PutValue(300);
        cells["A4"].PutValue("C"); cells["B4"].PutValue(2000); cells["C4"].PutValue(400);

        // Add a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIndex];

        // Add fields to the pivot table
        pivot.AddFieldToArea(PivotFieldType.Row, "Product");          // Row field
        pivot.AddFieldToArea(PivotFieldType.Data, "Revenue");        // Data field
        pivot.AddFieldToArea(PivotFieldType.Data, "Profit");         // Data field

        // Add a calculated field named "ProfitMargin" that computes Profit / Revenue
        // The third parameter 'true' drags the field to the data area automatically
        pivot.AddCalculatedField("ProfitMargin", "=Profit/Revenue", true);

        // Format the calculated field as a percentage
        PivotField profitMarginField = pivot.DataFields[pivot.DataFields.Count - 1];
        profitMarginField.NumberFormat = "0.00%";

        // Refresh and calculate the pivot table data
        pivot.RefreshData();
        pivot.CalculateData();

        // Save the workbook
        workbook.Save("PivotTable_With_ProfitMargin.xlsx");
    }
}
