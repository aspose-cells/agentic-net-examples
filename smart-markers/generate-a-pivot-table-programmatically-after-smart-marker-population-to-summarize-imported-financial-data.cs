// Title: Generate a Pivot Table After Smart‑Marker Population with Aspose.Cells for .NET
// Description: This Aspose.Cells for .NET sample builds a workbook template with smart markers for Region, Product, Month, and Revenue, populates it from a DataTable of financial records using WorkbookDesigner, then creates a separate worksheet and adds a pivot table that groups revenue by Region (rows) and Month (columns). The final Excel file demonstrates automated financial reporting in C#.
// Keywords: Aspose.Cells | smart markers | pivot table | .NET | C# | financial data | WorkbookDesigner | DataTable | Excel automation | region revenue pivot | programmatic pivot | Excel reporting
// Common Searches: Aspose.Cells create pivot table after smart markers | C# smart markers to pivot table example | generate financial pivot report with Aspose.Cells | populate worksheet with smart markers and add pivot | automate Excel pivot table using Aspose.Cells .NET
// Developer Intent: Build a pivot table that summarizes data filled via smart markers.
// Use Cases: Produce a monthly revenue summary by region and month without manual Excel work. | Design a reusable Excel template that auto‑populates sales data and creates a pivot report. | Integrate smart‑marker data import and pivot generation into an automated financial‑reporting pipeline.
// AI Prompts: Show how to format the pivot table revenue column as currency. | Add a total row for each region in the generated pivot table. | Create a chart that visualizes the pivot data directly after the table is built. | Explain how to enable automatic refresh of the pivot when the underlying smart‑marker data changes.

using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using AsposeRange = Aspose.Cells.Range;

// This Aspose.Cells for .NET sample builds a workbook template with smart markers for Region, Product, Month, and Revenue, populates it from a DataTable of financial records using WorkbookDesigner, then creates a separate worksheet and adds a pivot table that groups revenue by Region (rows) and Month (columns). The final Excel file demonstrates automated financial reporting in C#.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook that will act as the template containing smart markers
            Workbook workbook = new Workbook();
            Worksheet templateSheet = workbook.Worksheets[0];
            templateSheet.Name = "Template";

            // Define column headers
            templateSheet.Cells["A1"].PutValue("Region");
            templateSheet.Cells["B1"].PutValue("Product");
            templateSheet.Cells["C1"].PutValue("Month");
            templateSheet.Cells["D1"].PutValue("Revenue");

            // Insert smart markers for line‑by‑line data population
            templateSheet.Cells["A2"].PutValue("&=$Region");
            templateSheet.Cells["B2"].PutValue("&=$Product");
            templateSheet.Cells["C2"].PutValue("&=$Month");
            templateSheet.Cells["D2"].PutValue("&=$Revenue");

            // Mark the range that contains the smart markers
            templateSheet.Cells.CreateRange("A2:D2").Name = "_CellsSmartMarkers";

            // Prepare the financial data source
            DataTable financialData = new DataTable("FinancialData");
            financialData.Columns.Add("Region", typeof(string));
            financialData.Columns.Add("Product", typeof(string));
            financialData.Columns.Add("Month", typeof(string));
            financialData.Columns.Add("Revenue", typeof(double));

            financialData.Rows.Add("North", "Widget", "Jan", 12000);
            financialData.Rows.Add("North", "Widget", "Feb", 15000);
            financialData.Rows.Add("South", "Gadget", "Jan", 8000);
            financialData.Rows.Add("South", "Gadget", "Feb", 9500);
            financialData.Rows.Add("East", "Widget", "Jan", 11000);
            financialData.Rows.Add("East", "Gadget", "Feb", 13000);

            // Process the smart markers and populate the worksheet with the data
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(financialData);
            designer.Process();

            // Add a new worksheet that will host the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotReport");

            // Build the source data reference for the pivot table (including headers)
            AsposeRange sourceRange = templateSheet.Cells.MaxDisplayRange;
            string sourceData = $"=Template!{sourceRange.Address}";

            // Add the pivot table using the (sourceData, destCellName, tableName) overload
            int pivotIndex = pivotSheet.PivotTables.Add(sourceData, "A3", "FinancialPivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure the pivot fields: Region → rows, Month → columns, Revenue → data
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Month");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Revenue");

            // Optional layout and calculation
            pivotTable.ShowInTabularForm();
            pivotTable.CalculateData();

            // Save the final workbook
            workbook.Save("FinancialPivotReport.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
