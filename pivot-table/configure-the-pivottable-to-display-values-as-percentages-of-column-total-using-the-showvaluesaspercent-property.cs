// Title: Display sales values as percentage of column total in an Aspose.Cells PivotTable using C#
// AI Prompts: Generate C# code that creates a workbook, adds a pivot table, and configures the data field to show values as PercentageOfColumn with Aspose.Cells. | Write the method call that applies PivotFieldDataDisplayFormat.PercentageOfColumn to a pivot field and then refreshes and calculates the pivot table.
// Common Searches: Aspose.Cells C# pivot table show values as percent of column total | Set PivotField ShowValuesAs to PercentageOfColumn in .NET workbook | Refresh and calculate Aspose.Cells pivot table after changing display format | Example of using PivotFieldDataDisplayFormat.PercentageOfColumn with Aspose.Cells | Create pivot table with product rows and quarter columns using Aspose.Cells C#
// Tags: Aspose.Cells pivot table display format percentage | C# configure pivot field as column total percent | Aspose.Cells refresh and calculate pivot | Create pivot table from range Aspose.Cells C# | PivotFieldDataDisplayFormat PercentageOfColumn usage

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    // The sample creates a new workbook, fills it with product, quarter, and sales data, adds a pivot table with Product as rows, Quarter as columns, and Sales as values, then sets the Sales data field to display as a percentage of the column total using ShowValuesAs with PivotFieldDataDisplayFormat.PercentageOfColumn, refreshes the pivot cache, calculates the table, and saves the workbook.
    public class ShowValuesAsPercentageOfColumn
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                // Columns: Product, Quarter, Sales
                sheet.Cells["A1"].Value = "Product";
                sheet.Cells["B1"].Value = "Quarter";
                sheet.Cells["C1"].Value = "Sales";

                sheet.Cells["A2"].Value = "Apple";
                sheet.Cells["B2"].Value = "Q1";
                sheet.Cells["C2"].Value = 1200;

                sheet.Cells["A3"].Value = "Apple";
                sheet.Cells["B3"].Value = "Q2";
                sheet.Cells["C3"].Value = 1500;

                sheet.Cells["A4"].Value = "Orange";
                sheet.Cells["B4"].Value = "Q1";
                sheet.Cells["C4"].Value = 800;

                sheet.Cells["A5"].Value = "Orange";
                sheet.Cells["B5"].Value = "Q2";
                sheet.Cells["C5"].Value = 950;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add fields to the pivot table
                // Row field: Product
                int rowFieldIdx = pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                // Column field: Quarter
                int columnFieldIdx = pivotTable.AddFieldToArea(PivotFieldType.Column, "Quarter");
                // Data field: Sales
                int dataFieldIdx = pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Retrieve the column field and data field objects
                PivotField columnField = pivotTable.ColumnFields[columnFieldIdx];
                PivotField dataField = pivotTable.DataFields[dataFieldIdx];

                // Configure the data field to show values as percentage of column total
                dataField.ShowValuesAs(
                    PivotFieldDataDisplayFormat.PercentageOfColumn,
                    columnField.BaseIndex,
                    PivotItemPositionType.Next,
                    0);

                // Refresh the pivot cache and calculate the pivot table
                pivotTable.RefreshData();      // Correct method to refresh cache
                pivotTable.CalculateData();

                // Define output file path
                string outputPath = "Pivot_ShowValuesAs_PercentageOfColumn.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ShowValuesAsPercentageOfColumn.Run();
        }
    }
}
