// Title: Create and Refresh a Pivot Table from Smart Markers with Aspose.Cells for .NET (C#)
// Description: This example shows how to build an Excel workbook, place smart‑marker placeholders for Category and Amount, generate a pivot table on a separate sheet, feed a DataTable to WorkbookDesigner, expand the source range automatically, refresh the pivot, and save the file as SmartMarkerPivot.xlsx using Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# | Smart Markers | Pivot Table | Refresh Pivot | WorkbookDesigner | DataTable to Excel | dynamic source range | Excel automation | US developers | European .NET community
// Common Searches: Aspose.Cells create pivot table from smart markers C# | refresh pivot after WorkbookDesigner process | smart marker source range expansion example | generate Excel pivot using DataTable Aspose.Cells | C# code for dynamic pivot table with smart markers
// Developer Intent: Build a pivot table that updates automatically after smart‑marker data merging.
// Use Cases: Monthly sales report where new transaction rows are inserted via smart markers and the summary pivot updates instantly. | Financial dashboard that pulls quarterly figures from a DataTable and reflects changes without manual pivot adjustments. | Automated inventory analysis where product categories are added through smart markers and the pivot view expands accordingly.
// AI Prompts: Provide C# code to add multiple value fields to the pivot after processing smart markers. | Explain how the smart‑marker expansion modifies the pivot source range and how to verify the refresh. | Suggest robust error‑handling patterns for WorkbookDesigner.Process and pivot refresh in Aspose.Cells.

using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// This example shows how to build an Excel workbook, place smart‑marker placeholders for Category and Amount, generate a pivot table on a separate sheet, feed a DataTable to WorkbookDesigner, expand the source range automatically, refresh the pivot, and save the file as SmartMarkerPivot.xlsx using Aspose.Cells for C#.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and a worksheet that will hold the source data
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Insert header row
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Amount");

            // Insert a smart‑marker row – it will be replaced with the actual data rows
            dataSheet.Cells["A2"].PutValue("&=Data.Category");
            dataSheet.Cells["B2"].PutValue("&=Data.Amount");

            // Add a separate worksheet for the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

            // Define the initial source range (it will be automatically expanded after smart‑marker processing)
            string sourceRange = "Data!A1:B2";

            // Create the pivot table
            int pivotIndex = pivotSheet.PivotTables.Add(sourceRange, "A1", "MyPivot");
            PivotTable pivot = pivotSheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Prepare the data that will replace the smart markers
            DataTable dt = new DataTable("Data");
            dt.Columns.Add("Category", typeof(string));
            dt.Columns.Add("Amount", typeof(double));
            dt.Rows.Add("Fruit", 1200);
            dt.Rows.Add("Vegetable", 800);
            dt.Rows.Add("Beverage", 1500);

            // Process smart markers using WorkbookDesigner (replaces SmartMarkerProcessor)
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Data", dt);
            designer.Process();

            // Refresh the pivot table so it reflects the newly merged data
            pivotSheet.RefreshPivotTables();

            // Save the workbook
            workbook.Save("SmartMarkerPivot.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
