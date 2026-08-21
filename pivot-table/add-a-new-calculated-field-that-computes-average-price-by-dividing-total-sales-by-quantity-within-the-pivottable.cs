// Title: Create an Average Price Calculated Field (Sales/Quantity) in an Aspose.Cells PivotTable – C#
// Description: Learn how to use Aspose.Cells for .NET to build a PivotTable, add a calculated field named AveragePrice that divides total Sales by Quantity, format it to two decimal places, refresh the pivot cache, recalculate the data, and save the workbook.
// Keywords: Aspose.Cells calculated field | C# PivotTable average price | Aspose.Cells format number | Refresh pivot data Aspose.Cells | Add calculated field PivotTable C# | Aspose.Cells PivotTable example | Average price formula Sales/Quantity
// Common Searches: Aspose.Cells add calculated field to PivotTable | C# calculate average price in PivotTable using Aspose.Cells | format calculated field two decimal places Aspose.Cells | refresh pivot cache after adding calculated field Aspose.Cells | Aspose.Cells PivotTable average price example
// Developer Intent: Add a calculated field that divides total sales by quantity to display average price in a PivotTable using Aspose.Cells for .NET.
// Use Cases: Insert a calculated field called AveragePrice with the formula =Sales/Quantity into an existing PivotTable. | Apply the numeric format "#,##0.00" to the calculated field so results show two decimal places. | Refresh the pivot cache and recalculate the PivotTable after adding the calculated field. | Save the workbook to a file (e.g., PivotTable_AveragePrice.xlsx).
// AI Prompts: Generate C# code with Aspose.Cells that creates a PivotTable and adds a calculated field named AveragePrice (Sales divided by Quantity). | Show how to set a two‑decimal‑place number format for a calculated field in an Aspose.Cells PivotTable using C#. | Explain the steps to refresh pivot data and recalculate after adding a calculated field in Aspose.Cells for .NET.

using Aspose.Cells;
using Aspose.Cells.Pivot;

// Learn how to use Aspose.Cells for .NET to build a PivotTable, add a calculated field named AveragePrice that divides total Sales by Quantity, format it to two decimal places, refresh the pivot cache, recalculate the data, and save the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data: Product, Sales, Quantity
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["C1"].PutValue("Quantity");

        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(200);
        sheet.Cells["C2"].PutValue(10);

        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["C3"].PutValue(5);

        sheet.Cells["A4"].PutValue("Orange");
        sheet.Cells["B4"].PutValue(300);
        sheet.Cells["C4"].PutValue(15);

        // Add a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "SalesPivot");
        PivotTable pivot = sheet.PivotTables[pivotIndex];

        // Add fields to the pivot table
        pivot.AddFieldToArea(PivotFieldType.Row, "Product");          // Row field
        pivot.AddFieldToArea(PivotFieldType.Data, "Sales");          // Data field 1
        pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");       // Data field 2

        // Add a calculated field that computes average price = Sales / Quantity
        pivot.AddCalculatedField("AveragePrice", "=Sales/Quantity");

        // Format the calculated field (last added data field) to show two decimal places
        pivot.DataFields[pivot.DataFields.Count - 1].NumberFormat = "#,##0.00";

        // Refresh the pivot cache and calculate the pivot data
        pivot.RefreshData();
        pivot.CalculateData();

        // Save the workbook
        workbook.Save("PivotTable_AveragePrice.xlsx");
    }
}
