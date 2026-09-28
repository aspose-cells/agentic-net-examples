// Title: How to define a custom numeric range grouping for a Sales field in an Aspose.Cells pivot table using C#
// AI Prompts: Generate C# code that creates a workbook, inserts Region and Sales data, builds a pivot table, and groups the Sales row field into numeric ranges from 0 to 10000 with a 2000 interval using Aspose.Cells. | Write C# to access the PivotField after grouping and print its Start, End, and Interval values with Aspose.Cells.
// Common Searches: Aspose.Cells C# group pivot table row field by numeric range 0-10000 interval 2000 | How to set custom numeric grouping for sales amounts in an Aspose.Cells pivot table | Retrieve numeric group settings from a pivot field using Aspose.Cells for .NET | Create a pivot table with sales data and custom range groups in C# Aspose.Cells | Aspose.Cells GroupBy method numeric range example C#
// Tags: Aspose.Cells pivot numeric grouping | C# GroupBy numeric range Aspose.Cells | sales amount range groups pivot | pivot field group settings retrieval | export workbook with custom groups .xlsx

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsCustomNumericGrouping
{
    // The example builds a workbook, adds Region and Sales columns, creates a pivot table, places Sales as a row field and Region as a column field, then applies a numeric grouping from 0 to 10000 with a 2000 interval to the Sales field. After refreshing and calculating the pivot, it prints the grouping parameters (Start, End, Interval) and saves the file as CustomNumericGrouping.xlsx.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and add sample data
                Workbook workbook = new Workbook();
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";

                dataSheet.Cells["A1"].PutValue("Region");
                dataSheet.Cells["B1"].PutValue("Sales");

                string[] regions = { "North", "South", "East", "West", "North", "South", "East", "West" };
                double[] sales =   { 1200, 3400, 5600, 7800, 2100, 4300, 6500, 8700 };

                for (int i = 0; i < regions.Length; i++)
                {
                    dataSheet.Cells[i + 1, 0].PutValue(regions[i]);   // Region column
                    dataSheet.Cells[i + 1, 1].PutValue(sales[i]);    // Sales column
                }

                // Add a worksheet for the pivot table
                Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

                // Create a pivot table based on the data range
                int pivotIndex = pivotSheet.PivotTables.Add(
                    "Data!A1:B9",   // source range (including header)
                    "A3",           // destination cell in pivot sheet
                    "SalesPivot"); // pivot table name

                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Configure fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Sales");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Get the Sales row field and apply custom numeric grouping
                PivotField salesField = pivotTable.RowFields[0];
                salesField.GroupBy(0.0, 10000.0, 2000.0, true);

                // Refresh and calculate the pivot table
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Retrieve and display group settings if available
                var groupSettings = salesField.GroupSettings as PivotNumbericRangeGroupSettings;
                if (groupSettings != null)
                {
                    Console.WriteLine("Numeric grouping created:");
                    Console.WriteLine($"Start = {groupSettings.Start}");
                    Console.WriteLine($"End   = {groupSettings.End}");
                    Console.WriteLine($"Interval = {groupSettings.Interval}");
                }
                else
                {
                    Console.WriteLine("Group settings were not created.");
                }

                // Save the workbook
                workbook.Save("CustomNumericGrouping.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
