// Title: How to enable ShowValuesColumn in an Aspose.Cells PivotTable using C# to display each data field in its own column
// AI Prompts: Write C# code that creates a workbook, adds a pivot table, and sets PivotTable.ShowValuesColumn = true so each data field appears in a separate column. | Show the steps to modify an existing Aspose.Cells PivotTable in .NET to turn on the ShowValuesColumn flag and then refresh and calculate the pivot data. | Provide a complete example that saves the workbook after enabling ShowValuesColumn for a pivot table, including cache refresh and data calculation.
// Common Searches: Aspose.Cells C# set ShowValuesColumn on pivot table | display each data field in its own column in Aspose.Cells pivot table | programmatically turn on ShowValuesColumn property for PivotTable using .NET | refresh pivot cache after changing ShowValuesColumn Aspose.Cells example | save workbook with pivot table showing values in separate columns C#
// Tags: Aspose.Cells pivot table ShowValuesColumn | C# enable separate value columns in pivot table | Aspose.Cells refresh pivot cache after property change | save workbook as xlsx with pivot layout Aspose.Cells | add pivot table and configure fields Aspose.Cells C#

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    // Creates a workbook, fills sample data, adds a pivot table, sets PivotTable.ShowValuesColumn = true to place each data field in its own column, refreshes the pivot cache, calculates the data, and saves the file as PivotTable_ShowValuesColumn.xlsx.
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

                sheet.Cells["A2"].PutValue("Electronics");
                sheet.Cells["B2"].PutValue("Laptop");
                sheet.Cells["C2"].PutValue(1200);

                sheet.Cells["A3"].PutValue("Electronics");
                sheet.Cells["B3"].PutValue("Phone");
                sheet.Cells["C3"].PutValue(800);

                sheet.Cells["A4"].PutValue("Furniture");
                sheet.Cells["B4"].PutValue("Chair");
                sheet.Cells["C4"].PutValue(150);

                sheet.Cells["A5"].PutValue("Furniture");
                sheet.Cells["B5"].PutValue("Table");
                sheet.Cells["C5"].PutValue(300);

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table: Category as row, Product as column, Sales as data
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Show each data field in its own column
                pivotTable.ShowValuesRow = true;

                // Refresh the pivot cache and calculate the pivot table data
                pivotTable.RefreshData();      // Correct method to refresh cache
                pivotTable.CalculateData();

                // Save the workbook
                workbook.Save("PivotTable_ShowValuesColumn.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
