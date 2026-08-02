// Title: C# – Add an Average Price calculated field (Sales ÷ Quantity) to an Aspose.Cells PivotTable
// Description: Creates a workbook with product, sales, and quantity data, builds a PivotTable, adds Product as rows, Sales and Quantity as data fields, defines a calculated field named "AveragePrice" using the formula "=Sales/Quantity", refreshes the cache, recalculates, and saves the result as an Excel file.
// Keywords: Aspose.Cells | C# | .NET | PivotTable | calculated field | average price | sales divided by quantity | Excel automation | data analysis | RefreshData | CalculateData
// Common Searches: Aspose.Cells add calculated field to pivot table | C# average price pivot table Aspose.Cells | how to divide sales by quantity in Aspose.Cells pivot | Aspose.Cells refresh pivot after calculated field | create pivot table with custom formula using Aspose.Cells
// Developer Intent: Generate a PivotTable and insert a calculated field that returns Sales ÷ Quantity.
// Use Cases: Produce a sales report that shows total sales, total units, and unit price per product. | Build a live dashboard where the average price updates automatically when source data changes. | Export an Excel workbook with a pre‑calculated average price for downstream financial analysis.
// AI Prompts: Write C# code with Aspose.Cells to add a calculated field called 'AveragePrice' that computes Sales divided by Quantity in a PivotTable. | Explain how to refresh and recalculate a PivotTable after adding a custom calculated field using Aspose.Cells. | Suggest ways to handle division‑by‑zero scenarios for the AveragePrice calculated field in the generated workbook.

using Aspose.Cells;
using Aspose.Cells.Pivot;

// Creates a workbook with product, sales, and quantity data, builds a PivotTable, adds Product as rows, Sales and Quantity as data fields, defines a calculated field named "AveragePrice" using the formula "=Sales/Quantity", refreshes the cache, recalculates, and saves the result as an Excel file.
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
        pivot.AddFieldToArea(PivotFieldType.Row, "Product");
        pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
        pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");

        // Add a calculated field that computes average price (Sales / Quantity)
        pivot.AddCalculatedField("AveragePrice", "=Sales/Quantity", true);

        // Refresh the pivot cache and calculate the data
        pivot.RefreshData();
        pivot.CalculateData();

        // Save the workbook with the pivot table and calculated field
        workbook.Save("PivotTable_AveragePrice.xlsx");
    }
}
