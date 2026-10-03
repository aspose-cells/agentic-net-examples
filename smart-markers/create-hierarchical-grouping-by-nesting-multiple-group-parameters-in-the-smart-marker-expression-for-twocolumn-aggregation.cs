// Title: Generate a hierarchical Excel report with nested Group smart markers for Region and Product aggregation using Aspose.Cells for .NET (C#)
// AI Prompts: Create C# code that builds an Excel workbook, adds a data sheet, and uses Aspose.Cells WorkbookDesigner with {Group:Region} and {Group:Product} smart markers to produce a two‑level grouped report showing total sales per product. | Modify the smart‑marker layout to include a third grouping level (e.g., Year) and calculate the average sales for each year while keeping the existing Region‑Product hierarchy. | Replace the {Sum:Sales} aggregation with {Count:Sales} in the smart‑marker expression and adjust the report to display the number of sales entries for each product within each region.
// Common Searches: how to use nested Group smart markers in Aspose.Cells C# to group by region and product | Aspose.Cells WorkbookDesigner two level grouping with sum aggregation example | C# generate hierarchical Excel report with smart markers grouping and aggregation | Aspose.Cells smart marker expression for grouping and summing sales data | nested grouping smart markers Aspose.Cells .NET tutorial
// Tags: nested group smart markers Aspose.Cells | C# hierarchical Excel grouping with smart markers | sum aggregation smart marker expression | region product grouping report Aspose.Cells | two‑level grouping using smart markers

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The sample creates a workbook, populates a data sheet with Region, Product, and Sales columns, and defines a report sheet that uses {Group:Region}, a nested {Group:Product}, and {Sum:Sales} smart markers. WorkbookDesigner processes the markers with the provided data source, generating a hierarchical report saved as HierarchicalGrouping.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add a data worksheet
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Populate sample data (Region, Product, Sales)
            Cells cells = dataSheet.Cells;
            cells["A1"].PutValue("Region");
            cells["B1"].PutValue("Product");
            cells["C1"].PutValue("Sales");

            var sampleData = new List<object>
            {
                new { Region = "North", Product = "A", Sales = 100 },
                new { Region = "North", Product = "B", Sales = 150 },
                new { Region = "South", Product = "A", Sales = 200 },
                new { Region = "South", Product = "B", Sales = 250 },
                new { Region = "North", Product = "A", Sales = 120 },
                new { Region = "South", Product = "A", Sales = 180 }
            };

            for (int i = 0; i < sampleData.Count; i++)
            {
                dynamic row = sampleData[i];
                cells[i + 2, 0].PutValue(row.Region);
                cells[i + 2, 1].PutValue(row.Product);
                cells[i + 2, 2].PutValue(row.Sales);
            }

            // Add a worksheet that will contain the smart‑marker report
            Worksheet reportSheet = workbook.Worksheets[workbook.Worksheets.Add()];
            reportSheet.Name = "Report";

            // Header row for the report
            reportSheet.Cells["A1"].PutValue("Region");
            reportSheet.Cells["B1"].PutValue("Product");
            reportSheet.Cells["C1"].PutValue("Total Sales");

            // ---- Smart marker layout ----
            // Row 2: first level group (Region)
            reportSheet.Cells["A2"].PutValue("{Group:Region}");
            // Row 3: second level group (Product) under the current Region
            reportSheet.Cells["A3"].PutValue("{Group:Product}");
            // Row 3 column C: aggregated sum of Sales for the current Product
            reportSheet.Cells["C3"].PutValue("{Sum:Sales}");
            // --------------------------------

            // Process the smart markers using WorkbookDesigner
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            // Provide the data source with a name that matches the root object used in the markers
            designer.SetDataSource("Data", sampleData);
            designer.Process();

            // Save the resulting workbook
            workbook.Save("HierarchicalGrouping.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
