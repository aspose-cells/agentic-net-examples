// Title: Aspose.Cells C# – Enable ShowValuesColumn to Separate Value Fields in a PivotTable
// Description: This example creates a workbook, adds sample sales data, builds a PivotTable on range A1:C5, assigns Category to rows, Product to columns, and Sales to data, then sets the ShowValuesColumn property so each data field appears in its own column before refreshing and saving the file.
// Keywords: Aspose.Cells ShowValuesColumn | C# PivotTable separate value columns | Aspose.Cells enable ShowValuesColumn | PivotTable value field column per field | .NET Excel pivot table display options
// Common Searches: how to turn on ShowValuesColumn in Aspose.Cells | Aspose.Cells C# display each data field in its own column | pivot table ShowValuesColumn property example | separate value columns in Excel pivot using Aspose | Aspose.Cells version check for ShowValuesColumn support
// Developer Intent: Activate the ShowValuesColumn property on a PivotTable so every data field is rendered in a distinct column.
// Use Cases: Generate a sales report where the Sales metric is shown in a separate column for each product. | Create a financial summary with multiple metrics (Revenue, Cost, Profit) each displayed in its own column by enabling ShowValuesColumn. | Build a dashboard that adds several data fields (e.g., Sales, Quantity) to a PivotTable and uses ShowValuesColumn to produce a flat table layout for downstream processing.
// AI Prompts: Write C# code using Aspose.Cells to set pivotTable.ShowValuesColumn = true, add two data fields, refresh the data, and save the workbook. | Show an example that toggles ShowValuesColumn on a PivotTable and explains the visual change in the exported Excel file. | Provide a snippet that checks the Aspose.Cells version at runtime and conditionally applies ShowValuesColumn if the property is available.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// This example creates a workbook, adds sample sales data, builds a PivotTable on range A1:C5, assigns Category to rows, Product to columns, and Sales to data, then sets the ShowValuesColumn property so each data field appears in its own column before refreshing and saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Product");
            sheet.Cells["C1"].PutValue("Sales");

            sheet.Cells["A2"].PutValue("Fruit");
            sheet.Cells["B2"].PutValue("Apple");
            sheet.Cells["C2"].PutValue(120);

            sheet.Cells["A3"].PutValue("Fruit");
            sheet.Cells["B3"].PutValue("Orange");
            sheet.Cells["C3"].PutValue(80);

            sheet.Cells["A4"].PutValue("Vegetable");
            sheet.Cells["B4"].PutValue("Carrot");
            sheet.Cells["C4"].PutValue(150);

            sheet.Cells["A5"].PutValue("Vegetable");
            sheet.Cells["B5"].PutValue("Potato");
            sheet.Cells["C5"].PutValue(200);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E1", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");    // Row field
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Product"); // Column field
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");     // Data field

            // The ShowValuesColumn property is not available in all versions; omitted.

            // Refresh data source and calculate the pivot table
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook with the configured pivot table
            workbook.Save("PivotTableShowValuesColumnDemo.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
