// Title: Clone a PivotTable, modify its data source range, and insert it into a new worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that clones an existing PivotTable, updates its source data range with ChangeDataSource, refreshes and recalculates the pivot, and saves the workbook using Aspose.Cells. | Provide a method that copies a PivotTable to another worksheet, changes the data source to a different range, and writes the result to an XLSX file. | Show how to use PivotTableCollection.Add to duplicate a pivot, call ChangeDataSource, then invoke RefreshData and CalculateData in Aspose.Cells.
// Common Searches: Aspose.Cells C# clone pivot table and set a new source range on a different sheet | How to duplicate a PivotTable and change its data source with Aspose.Cells for .NET | C# example for copying a pivot, updating source range, and saving as XLSX using Aspose.Cells
// Tags: clone pivot table Aspose.Cells C# | change pivot data source Aspose.Cells | add cloned pivot to new worksheet | refresh pivot after source change Aspose.Cells | save workbook as XLSX Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates cloning an existing PivotTable, updating its data source range with ChangeDataSource, refreshing and recalculating the pivot, and saving the workbook as an XLSX file using Aspose.Cells for .NET.
    public class ClonePivotTableModifySourceDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sourceSheet = workbook.Worksheets[0];
                sourceSheet.Name = "SourceData";

                // Populate sample data for the original pivot table
                sourceSheet.Cells["A1"].PutValue("Category");
                sourceSheet.Cells["B1"].PutValue("Product");
                sourceSheet.Cells["C1"].PutValue("Sales");

                sourceSheet.Cells["A2"].PutValue("Fruit");
                sourceSheet.Cells["B2"].PutValue("Apple");
                sourceSheet.Cells["C2"].PutValue(120);

                sourceSheet.Cells["A3"].PutValue("Fruit");
                sourceSheet.Cells["B3"].PutValue("Banana");
                sourceSheet.Cells["C3"].PutValue(80);

                sourceSheet.Cells["A4"].PutValue("Vegetable");
                sourceSheet.Cells["B4"].PutValue("Carrot");
                sourceSheet.Cells["C4"].PutValue(150);

                // Define the source data range for the pivot table
                string sourceData = "=SourceData!A1:C4";

                // Add the original pivot table to the source sheet
                PivotTableCollection sourcePivots = sourceSheet.PivotTables;
                int sourcePivotIndex = sourcePivots.Add(sourceData, "E1", "OriginalPivot");
                PivotTable sourcePivot = sourcePivots[sourcePivotIndex];

                // Configure the original pivot table
                sourcePivot.AddFieldToArea(PivotFieldType.Row, "Category");
                sourcePivot.AddFieldToArea(PivotFieldType.Row, "Product");
                sourcePivot.AddFieldToArea(PivotFieldType.Data, "Sales");
                sourcePivot.CalculateData();

                // Add a new worksheet where the cloned pivot table will be placed
                Worksheet targetSheet = workbook.Worksheets.Add("ClonedPivotSheet");

                // Clone the original pivot table into the new worksheet at cell A1
                PivotTableCollection targetPivots = targetSheet.PivotTables;
                int clonedPivotIndex = targetPivots.Add(sourcePivot, "A1", "ClonedPivot");
                PivotTable clonedPivot = targetPivots[clonedPivotIndex];

                // Modify the data source of the cloned pivot table to a different range
                // Here we use a new range within the original sheet (A1:C4) – you can adjust as needed
                string[] newDataSource = new string[] { "A1:C4", "SourceData" };
                clonedPivot.ChangeDataSource(newDataSource);

                // Refresh the pivot table to reflect the new data source
                clonedPivot.RefreshData();

                // Recalculate the cloned pivot table
                clonedPivot.CalculateData();

                // Save the workbook to a file
                workbook.Save("ClonedPivotWithNewSource.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point required for console application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
