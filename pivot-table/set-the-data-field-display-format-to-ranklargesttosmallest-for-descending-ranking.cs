// Title: Aspose.Cells for .NET – Set Pivot Table Data Field to Rank Largest‑to‑Smallest (Descending)
// Description: This C# example creates a workbook, adds sample sales data, builds a pivot table, and configures the Quantity data field to use the RankLargestToSmallest display format. The pivot calculates ranks from highest to lowest before saving the workbook as an Excel file.
// Keywords: Aspose.Cells | C# | .NET | pivot table rank largest to smallest | descending ranking | PivotFieldDataDisplayFormat | ShowValuesSetting | Excel automation | data field ranking | Aspose.Cells example
// Common Searches: Aspose.Cells set pivot table rank descending | C# rank largest to smallest pivot field | ShowValuesSetting calculation type example | How to rank pivot table values in Aspose.Cells | PivotFieldDataDisplayFormat RankLargestToSmallest usage
// Developer Intent: Apply a descending rank display format to a pivot table data field using Aspose.Cells for .NET.
// Use Cases: Sales dashboard that lists products from highest to lowest quantity sold. | Inventory report ranking categories by total stock in descending order. | Employee performance sheet showing staff ranked by achievement scores.
// AI Prompts: Generate C# code that sets a pivot table data field to RankLargestToSmallest with Aspose.Cells. | Explain the effect of ShowValuesSetting.CalculationType on pivot table ranking in Aspose.Cells. | Adapt the example to apply RankLargestToSmallest to multiple data fields in the same pivot table.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // This C# example creates a workbook, adds sample sales data, builds a pivot table, and configures the Quantity data field to use the RankLargestToSmallest display format. The pivot calculates ranks from highest to lowest before saving the workbook as an Excel file.
    public class SetRankLargestToSmallestDemo
    {
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
                cells["C1"].Value = "Quantity";

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
                cells["C6"].Value = 10;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:C6", "E3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add a row field (Category) and a data field (Quantity)
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");

                // Retrieve the first data field and set ranking to largest‑to‑smallest
                PivotField dataField = pivotTable.DataFields[0];
                dataField.ShowValuesSetting.CalculationType = PivotFieldDataDisplayFormat.RankLargestToSmallest;

                // Calculate the pivot data
                pivotTable.CalculateData();

                // Save the workbook
                string outputPath = "SetRankLargestToSmallestDemo_out.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetRankLargestToSmallestDemo.Run();
        }
    }
}
