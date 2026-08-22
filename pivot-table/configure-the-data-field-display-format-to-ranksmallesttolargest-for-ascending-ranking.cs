// Title: How to set a PivotTable data field to RankSmallestToLargest (ascending ranking) using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a PivotTable, adds row, column, and data fields, and sets the data field's ShowValuesSetting.CalculationType to RankSmallestToLargest. | Provide a step‑by‑step example of refreshing the pivot cache after applying the RankSmallestToLargest display format with Aspose.Cells. | Show how to save the workbook after configuring the pivot field ranking to smallest‑to‑largest in a .xlsx file.
// Common Searches: Aspose.Cells C# set pivot table data field ranking to smallest to largest | How to apply RankSmallestToLargest calculation type to a PivotField in .NET | Example of using ShowValuesSetting.CalculationType for ascending ranking in Aspose.Cells | C# code to create a pivot table and rank values from lowest to highest with Aspose.Cells
// Tags: Aspose.Cells pivot field rank ascending | set ShowValuesSetting.CalculationType RankSmallestToLargest | C# create pivot table Aspose.Cells | Aspose.Cells refresh pivot cache | save workbook as .xlsx Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // // Demonstrates creating a workbook, populating sample data, adding a PivotTable, setting the data field's display format to RankSmallestToLargest (ascending ranking), refreshing the pivot cache, calculating data, and saving the result as an .xlsx file.
    public class PivotFieldRankSmallestToLargestDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data for the pivot table
                cells["A1"].Value = "Category";
                cells["B1"].Value = "Item";
                cells["C1"].Value = "Amount";

                cells["A2"].Value = "Fruit";
                cells["B2"].Value = "Apple";
                cells["C2"].Value = 30;

                cells["A3"].Value = "Fruit";
                cells["B3"].Value = "Banana";
                cells["C3"].Value = 20;

                cells["A4"].Value = "Fruit";
                cells["B4"].Value = "Orange";
                cells["C4"].Value = 50;

                cells["A5"].Value = "Vegetable";
                cells["B5"].Value = "Carrot";
                cells["C5"].Value = 40;

                cells["A6"].Value = "Vegetable";
                cells["B6"].Value = "Broccoli";
                cells["C6"].Value = 25;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:C6", "E3", "SalesPivot");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Add fields to the pivot table
                pivot.AddFieldToArea(PivotFieldType.Row, "Item");          // Row field
                pivot.AddFieldToArea(PivotFieldType.Column, "Category");  // Column field
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");      // Data field

                // Get the data field that was just added
                PivotField dataField = pivot.DataFields[0];

                // Set the display format to rank smallest to largest (ascending ranking)
                dataField.ShowValuesSetting.CalculationType = PivotFieldDataDisplayFormat.RankSmallestToLargest;

                // Refresh pivot cache and calculate the pivot table data
                pivot.RefreshData();
                pivot.CalculateData();

                // Save the workbook to a file
                workbook.Save("PivotField_RankSmallestToLargest.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Run error: {ex.Message}");
            }
        }
    }
}
