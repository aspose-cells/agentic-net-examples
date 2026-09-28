// Title: How to keep two Aspose.Cells slicers synchronized when selecting an item in C#
// AI Prompts: Write C# code that selects a specific value in one Aspose.Cells slicer and automatically updates another slicer linked to the same pivot field. | Show how to refresh multiple slicers after modifying the slicer cache so their selections stay in sync in an Excel workbook using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# synchronize slicer selections across multiple slicers | refresh second slicer after changing first slicer in Aspose.Cells | programmatically select slicer item and update linked slicer Aspose.Cells .NET | how to link two slicers to the same pivot field using Aspose.Cells | C# example of slicer cache item selection with Aspose.Cells
// Tags: Aspose.Cells slicer cache selection | C# refresh multiple slicers | synchronizing slicers pivot table Aspose.Cells | programmatic slicer synchronization .NET | Excel workbook slicer linking Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;      // Pivot table related classes
using Aspose.Cells.Slicers;    // Slicer related classes

namespace AsposeCellsExamples
{
    // The example creates a workbook, adds sample data, builds a pivot table, inserts two slicers on the 'Fruit' field, selects 'Apple' in the first slicer, refreshes both slicers to keep them synchronized, and saves the file as SynchronizedSlicers.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data
                cells["A1"].Value = "Fruit";
                cells["B1"].Value = "Year";
                cells["C1"].Value = "Amount";

                string[] fruits = { "Apple", "Banana", "Cherry", "Apple", "Banana", "Cherry" };
                int[] years = { 2020, 2020, 2020, 2021, 2021, 2021 };
                int[] amounts = { 50, 70, 90, 55, 75, 95 };

                for (int i = 0; i < fruits.Length; i++)
                {
                    cells[i + 1, 0].Value = fruits[i];
                    cells[i + 1, 1].Value = years[i];
                    cells[i + 1, 2].Value = amounts[i];
                }

                // Add a pivot table based on the data range A1:C7
                PivotTableCollection pivots = sheet.PivotTables;
                int pivotIdx = pivots.Add("A1:C7", "E1", "FruitPivot");
                PivotTable pivot = pivots[pivotIdx];

                // Configure pivot fields
                pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
                pivot.AddFieldToArea(PivotFieldType.Column, "Year");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Apply style and refresh
                pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
                pivot.RefreshData();
                pivot.CalculateData();

                // Add two slicers that target the same pivot field ("Fruit")
                SlicerCollection slicers = sheet.Slicers;
                int slicerIdx1 = slicers.Add(pivot, "G1", "Fruit");
                int slicerIdx2 = slicers.Add(pivot, "G10", "Fruit");

                Slicer slicer1 = slicers[slicerIdx1];
                Slicer slicer2 = slicers[slicerIdx2];

                // ---- Synchronization logic ----
                // Example: select the first item ("Apple") in slicer1
                foreach (SlicerCacheItem item in slicer1.SlicerCache.SlicerCacheItems)
                {
                    // Select only the item whose value is "Apple"
                    item.Selected = item.Value?.ToString() == "Apple";
                }

                // Refresh both slicers so that the second slicer reflects the same selection
                slicer1.Refresh();   // updates the pivot table based on slicer1 selection
                slicer2.Refresh();   // reads the pivot table state and updates its own cache

                // Save the workbook
                string outputPath = "SynchronizedSlicers.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
