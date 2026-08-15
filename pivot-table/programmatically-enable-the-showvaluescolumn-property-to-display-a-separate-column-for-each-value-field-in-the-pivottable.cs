// Title: Aspose.Cells C# – ShowValuesColumn property for PivotTable (unsupported in current version)
// Description: C# example that creates a workbook, fills a small sales table, adds a PivotTable and assigns row, column, and data fields. The code notes that the ShowValuesColumn property, which would place each value field in its own column, is not available in the present Aspose.Cells release. The workbook is then calculated and saved as PivotTable_ShowValuesColumn.xlsx.
// Keywords: Aspose.Cells | ShowValuesColumn | PivotTable layout | C# | .NET | Excel pivot value columns | value field separate column | Aspose.Cells version compatibility | pivot table programming | workaround for ShowValuesColumn
// Common Searches: Aspose.Cells ShowValuesColumn support | how to display each value field in its own column with Aspose.Cells | pivot table value column layout .NET | Aspose.Cells version that includes ShowValuesColumn | alternative to ShowValuesColumn in Aspose.Cells
// Developer Intent: Attempt to activate the ShowValuesColumn setting so that each data field appears in a distinct column of the generated PivotTable using Aspose.Cells for .NET.
// Use Cases: Building a sales dashboard where Sales, Quantity, and Discount are shown in separate columns for easy comparison. | Creating a financial report that lists revenue, cost, and profit as individual columns in a PivotTable. | Exporting pivot data to Excel for downstream BI tools that require one column per metric.
// AI Prompts: Generate C# code with Aspose.Cells that checks the library version and enables ShowValuesColumn if the property exists. | Suggest a workaround to simulate ShowValuesColumn behavior when the property is missing in Aspose.Cells. | Explain how to restructure a PivotTable in Aspose.Cells to display each value field in its own column without using ShowValuesColumn.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    // C# example that creates a workbook, fills a small sales table, adds a PivotTable and assigns row, column, and data fields. The code notes that the ShowValuesColumn property, which would place each value field in its own column, is not available in the present Aspose.Cells release. The workbook is then calculated and saved as PivotTable_ShowValuesColumn.xlsx.
    class ShowValuesColumnDemo
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

                sheet.Cells["A2"].PutValue("Electronics");
                sheet.Cells["B2"].PutValue("Laptop");
                sheet.Cells["C2"].PutValue(1200);

                sheet.Cells["A3"].PutValue("Electronics");
                sheet.Cells["B3"].PutValue("Phone");
                sheet.Cells["C3"].PutValue(800);

                sheet.Cells["A4"].PutValue("Furniture");
                sheet.Cells["B4"].PutValue("Chair");
                sheet.Cells["C4"].PutValue(150);

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:C4", "E1", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add fields to the pivot table
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // NOTE: The ShowValuesColumn property is not available in this version of Aspose.Cells.
                // The pivot table will use the default layout for value fields.

                // Calculate the pivot table data
                pivotTable.CalculateData();

                // Save the workbook
                workbook.Save("PivotTable_ShowValuesColumn.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
