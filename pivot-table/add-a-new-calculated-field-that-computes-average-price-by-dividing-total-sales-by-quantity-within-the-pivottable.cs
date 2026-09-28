// Title: Create a PivotTable with an Average Price calculated field (Sales ÷ Quantity) using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that builds a PivotTable with Aspose.Cells, adds Sales and Quantity as data fields, and defines a calculated field named AveragePrice using the formula =Sales/Quantity. | Demonstrate how to refresh the PivotTable cache and recalculate its data after inserting a custom calculated field in Aspose.Cells.
// Common Searches: how to define a calculated field in an Aspose.Cells PivotTable C# | Aspose.Cells example for average price = sales / quantity in pivot | C# code to add custom formula to PivotTable using Aspose.Cells library | refresh pivot data after adding calculated field Aspose.Cells .NET | create pivot table with sales and quantity fields and compute average price Aspose.Cells
// Tags: add calculated field to Aspose.Cells PivotTable | average price calculation in Aspose.Cells pivot | refresh pivot data Aspose.Cells C# | Aspose.Cells PivotTable custom formula | C# Aspose.Cells sales quantity pivot example

using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample creates a workbook, populates it with product, sales, and quantity data, builds a PivotTable on that range, adds Sales and Quantity as data fields, inserts a calculated field named AveragePrice using the formula =Sales/Quantity, refreshes and calculates the pivot data, and saves the workbook as an XLSX file.
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

        // Configure pivot fields
        pivot.AddFieldToArea(PivotFieldType.Row, "Product");      // Row field
        pivot.AddFieldToArea(PivotFieldType.Data, "Sales");      // Data field
        pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");   // Data field

        // Add a calculated field that computes average price = Sales / Quantity
        pivot.AddCalculatedField("AveragePrice", "=Sales/Quantity", true);

        // Refresh the pivot cache and calculate the pivot data
        pivot.RefreshData();
        pivot.CalculateData();

        // Save the workbook with the pivot table and calculated field
        workbook.Save("PivotTableWithAveragePrice.xlsx");
    }
}
