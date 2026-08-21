// Title: C# – Add a Year‑Over‑Year Growth Calculated Field to an Aspose.Cells Pivot Table
// Description: This Aspose.Cells for .NET example creates a workbook with product sales for 2022‑2023, builds a pivot table (Product rows, Year columns, Sales sum), adds a calculated field named YoYGrowth, configures it to show the percentage difference from the previous year, formats the result as a two‑decimal percent, refreshes the cache, calculates the data, and saves the file as PivotTable_YearOverYearGrowth.xlsx.
// Keywords: Aspose.Cells C# pivot table | add calculated field Aspose | year over year growth Excel | YoY percentage difference | ShowValuesAs PivotField | PivotFieldDataDisplayFormat percentage | Excel sales report automation | GitHub Aspose.Cells example | dynamic pivot cache refresh | C# Excel export
// Common Searches: how to add YoY growth calculated field in Aspose.Cells | Aspose.Cells pivot table show values as percentage difference | C# create pivot table with calculated field | format pivot field as percent Aspose.Cells | refresh pivot cache after adding calculated field
// Developer Intent: Generate a pivot table in C# and insert a calculated field that displays each product’s year‑over‑year sales growth as a formatted percentage.
// Use Cases: Produce a quarterly sales dashboard that automatically calculates YoY growth for every product. | Export dynamic Excel reports where new sales data is added and the pivot table updates the growth percentages without manual edits. | Integrate automated financial analysis into a .NET application, showing percentage change between consecutive years.
// AI Prompts: Write C# code using Aspose.Cells to add a calculated field called 'YoYGrowth' that shows the percentage difference from the previous year in a pivot table. | Explain the steps to configure ShowValuesAs with PivotFieldDataDisplayFormat.PercentageDifferenceFrom for a pivot field in Aspose.Cells. | Provide a snippet that formats a pivot field as a percentage with two decimal places and refreshes the pivot cache after adding a calculated field.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // This Aspose.Cells for .NET example creates a workbook with product sales for 2022‑2023, builds a pivot table (Product rows, Year columns, Sales sum), adds a calculated field named YoYGrowth, configures it to show the percentage difference from the previous year, formats the result as a two‑decimal percent, refreshes the cache, calculates the data, and saves the file as PivotTable_YearOverYearGrowth.xlsx.
    public class PivotTableYearOverYearGrowth
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data: Product, Year, Sales
                cells["A1"].Value = "Product";
                cells["B1"].Value = "Year";
                cells["C1"].Value = "Sales";

                cells["A2"].Value = "Apple";   cells["B2"].Value = 2022; cells["C2"].Value = 1200;
                cells["A3"].Value = "Apple";   cells["B3"].Value = 2023; cells["C3"].Value = 1500;
                cells["A4"].Value = "Banana";  cells["B4"].Value = 2022; cells["C4"].Value = 800;
                cells["A5"].Value = "Banana";  cells["B5"].Value = 2023; cells["C5"].Value = 950;
                cells["A6"].Value = "Cherry";  cells["B6"].Value = 2022; cells["C6"].Value = 500;
                cells["A7"].Value = "Cherry";  cells["B7"].Value = 2023; cells["C7"].Value = 650;

                // Create a pivot table covering the data range
                int pivotIndex = sheet.PivotTables.Add("A1:C7", "E3", "SalesPivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure pivot layout: rows = Product, columns = Year, data = Sales (sum)
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Year");
                int salesDataFieldIdx = pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");
                PivotField salesDataField = pivotTable.DataFields[salesDataFieldIdx];

                // Add a calculated field that references Sales; YoY will be shown via ShowValuesAs
                pivotTable.AddCalculatedField("YoYGrowth", "=Sales", true);
                PivotField yoyField = pivotTable.DataFields[pivotTable.DataFields.Count - 1];

                // Show YoY as percentage difference from the previous year
                PivotField yearField = pivotTable.ColumnFields[0];
                int baseFieldIndex = yearField.BaseIndex;
                yoyField.ShowValuesAs(
                    PivotFieldDataDisplayFormat.PercentageDifferenceFrom,
                    baseFieldIndex,
                    PivotItemPositionType.Previous,
                    0);

                // Format as percentage with two decimal places
                yoyField.NumberFormat = "0.00%";

                // Refresh pivot cache and calculate data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                string outputPath = "PivotTable_YearOverYearGrowth.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for console application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
