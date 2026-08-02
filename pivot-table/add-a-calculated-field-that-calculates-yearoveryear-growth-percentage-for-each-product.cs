// Title: Create a YoY Growth Calculated Field in an Aspose.Cells Pivot Table (C#)
// Description: Shows how to build a workbook with product sales data, generate a pivot table (Product rows, Year columns, Sales values), add a calculated field named YoY_Growth that computes (Sales‑PrevYearSales)/PrevYearSales, format it as a percentage, refresh the pivot, and save the result as an .xlsx file.
// Keywords: Aspose.Cells | C# pivot table | calculated field | year over year growth | YoY growth percentage | Excel pivot | Aspose.Cells .NET | pivot refresh | percentage format | sales analysis
// Common Searches: Aspose.Cells add calculated field to pivot | C# YoY growth pivot table | calculate year over year sales Aspose.Cells | format pivot field as percentage Aspose.Cells | refresh pivot after calculated field C#
// Developer Intent: Add a percentage‑formatted year‑over‑year growth calculated field to a pivot table using Aspose.Cells for .NET.
// Use Cases: Generate a sales dashboard that automatically displays growth between consecutive years for each product. | Create reusable Excel reports where new data updates recalculate YoY percentages without manual intervention. | Export financial or inventory data with built‑in growth metrics directly inside the pivot view.
// AI Prompts: Write C# code with Aspose.Cells to add a YoY_Growth calculated field to a pivot table and format it as 0.00%. | Explain how to refresh and recalculate a pivot table after inserting a calculated field in Aspose.Cells. | Show how to prevent division‑by‑zero errors in a YoY growth formula within an Aspose.Cells calculated field.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Shows how to build a workbook with product sales data, generate a pivot table (Product rows, Year columns, Sales values), add a calculated field named YoY_Growth that computes (Sales‑PrevYearSales)/PrevYearSales, format it as a percentage, refresh the pivot, and save the result as an .xlsx file.
    public class PivotTableYearOverYearGrowth
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data: Product, Year, Sales
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Year");
                sheet.Cells["C1"].PutValue("Sales");

                // Sample rows
                sheet.Cells["A2"].PutValue("Apple");   sheet.Cells["B2"].PutValue(2020); sheet.Cells["C2"].PutValue(1200);
                sheet.Cells["A3"].PutValue("Apple");   sheet.Cells["B3"].PutValue(2021); sheet.Cells["C3"].PutValue(1500);
                sheet.Cells["A4"].PutValue("Orange");  sheet.Cells["B4"].PutValue(2020); sheet.Cells["C4"].PutValue(800);
                sheet.Cells["A5"].PutValue("Orange");  sheet.Cells["B5"].PutValue(2021); sheet.Cells["C5"].PutValue(950);
                sheet.Cells["A6"].PutValue("Banana");  sheet.Cells["B6"].PutValue(2020); sheet.Cells["C6"].PutValue(500);
                sheet.Cells["A7"].PutValue("Banana");  sheet.Cells["B7"].PutValue(2021); sheet.Cells["C7"].PutValue(650);

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:C7", "E3", "SalesPivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot layout: Product in rows, Year in columns, Sales as data
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Year");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Add a calculated field that computes Year‑over‑Year growth percentage
                // Formula: (Sales - PrevYearSales) / PrevYearSales
                string growthFormula = "=(Sales - PrevYearSales) / PrevYearSales";
                pivotTable.AddCalculatedField("YoY_Growth", growthFormula, true);

                // Format the newly added calculated field as a percentage with two decimal places
                PivotField growthField = pivotTable.DataFields[pivotTable.DataFields.Count - 1];
                growthField.NumberFormat = "0.00%";

                // Refresh and calculate the pivot table to apply the changes
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Determine output file path
                string outputPath = "PivotTable_YearOverYearGrowth.xlsx";

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
            PivotTableYearOverYearGrowth.Run();
        }
    }
}
