// Title: Refresh a PivotTable After Modifying Source Data with Aspose.Cells (C#)
// Description: Learn how to programmatically refresh an Aspose.Cells PivotTable after changing source worksheet values using RefreshData and CalculateData, then save the updated workbook.
// Keywords: Aspose.Cells | C# | PivotTable refresh | RefreshData | CalculateData | update pivot cache | Excel automation | .NET Excel library | dynamic report generation | programmatic pivot update
// Common Searches: Aspose.Cells refresh pivot table after data change C# | How to update pivot cache programmatically in Aspose.Cells | C# code to recalculate PivotTable without reopening workbook | RefreshData vs CalculateData Aspose.Cells PivotTable | Dynamic Excel report with refreshed PivotTable using Aspose
// Developer Intent: Update an existing PivotTable so it reflects edited source data without recreating the table.
// Use Cases: After adjusting sales figures in the source sheet, call RefreshData and CalculateData to keep the pivot report accurate. | In automated reporting pipelines, refresh the pivot before exporting to ensure the latest data is displayed. | When batch‑processing large datasets, invoke the refresh methods after each batch to maintain a current PivotTable view.
// AI Prompts: Show C# code that refreshes an Aspose.Cells PivotTable after source cells are modified. | Explain the roles of RefreshData and CalculateData when updating a PivotTable with Aspose.Cells. | Provide a step‑by‑step guide to programmatically recalculate a PivotTable without recreating it in C#.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Learn how to programmatically refresh an Aspose.Cells PivotTable after changing source worksheet values using RefreshData and CalculateData, then save the updated workbook.
    public class RefreshPivotTableAfterDataChange
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // Get the first worksheet where source data resides
                Worksheet dataSheet = workbook.Worksheets[0];

                // Populate sample source data for the pivot table
                dataSheet.Cells["A1"].PutValue("Product");
                dataSheet.Cells["B1"].PutValue("Sales");
                dataSheet.Cells["A2"].PutValue("Apple");
                dataSheet.Cells["B2"].PutValue(100);
                dataSheet.Cells["A3"].PutValue("Banana");
                dataSheet.Cells["B3"].PutValue(200);
                dataSheet.Cells["A4"].PutValue("Apple");
                dataSheet.Cells["B4"].PutValue(150);

                // Add a second worksheet to host the pivot table
                Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

                // Create the pivot table based on the source range A1:B4
                int pivotIndex = pivotSheet.PivotTables.Add("A1:B4", "E3", "SalesPivot");
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Configure the pivot fields (Product as row, Sales as data)
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Column 0 -> Product
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Column 1 -> Sales

                // Initial calculation so the pivot table shows data
                pivotTable.RefreshData();      // Refresh cache
                pivotTable.CalculateData();    // Recalculate pivot

                // ----- Modify the source data -----
                dataSheet.Cells["B2"].PutValue(120); // Change Apple sales from 100 to 120
                dataSheet.Cells["B3"].PutValue(250); // Change Banana sales from 200 to 250

                // Refresh the pivot table to reflect the updated source data
                pivotTable.RefreshData();      // Refresh cache after data change
                pivotTable.CalculateData();    // Recalculate pivot

                // Save the workbook with the refreshed pivot table
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RefreshedPivotTable.xlsx");
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
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
            RefreshPivotTableAfterDataChange.Run();
        }
    }
}
