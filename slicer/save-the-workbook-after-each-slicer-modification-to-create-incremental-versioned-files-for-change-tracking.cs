// Title: Incrementally Save Excel Workbook After Each Slicer Change Using Aspose.Cells for .NET
// Description: C# sample that creates a workbook, adds a pivot table and a linked slicer, then modifies slicer properties (caption, lock, column count, size) and saves a new file after each change (SlicerVersion_0.xlsx – SlicerVersion_5.xlsx) to provide versioned change tracking.
// Keywords: Aspose.Cells | C# slicer | Excel slicer versioning | incremental workbook save | pivot table slicer | track slicer changes | Workbook.Save | .NET Excel automation | change log Excel | slicer property refresh
// Common Searches: how to save Excel after each slicer modification Aspose.Cells | C# versioned files for slicer property changes | track slicer adjustments with incremental saves | Aspose.Cells example for slicer change tracking | save workbook multiple times after slicer refresh
// Developer Intent: Create a series of sequentially named Excel files that capture every slicer property update for audit and change‑tracking purposes.
// Use Cases: Generate an initial workbook, then record each caption change as a separate versioned file. | Lock or unlock slicer position, refresh the view, and persist each state for compliance reporting. | Alter slicer layout (columns, width, height), refresh the pivot table, and save distinct files to compare visual configurations.
// AI Prompts: Write C# code with Aspose.Cells that adds a slicer to a pivot table and saves the workbook after each slicer property change using sequential filenames. | Provide a helper method that accepts a slicer, applies a given property change, refreshes it, and saves the workbook with an incremented version number. | Explain how to implement change tracking for slicer modifications by creating incremental workbook versions in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerVersioningDemo
{
    // C# sample that creates a workbook, adds a pivot table and a linked slicer, then modifies slicer properties (caption, lock, column count, size) and saves a new file after each change (SlicerVersion_0.xlsx – SlicerVersion_5.xlsx) to provide versioned change tracking.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Prepare sample data for a pivot table
            // -------------------------------------------------
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Header row
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Amount");

            // Sample rows
            dataSheet.Cells["A2"].PutValue("Food");
            dataSheet.Cells["B2"].PutValue(120);
            dataSheet.Cells["A3"].PutValue("Food");
            dataSheet.Cells["B3"].PutValue(80);
            dataSheet.Cells["A4"].PutValue("Beverage");
            dataSheet.Cells["B4"].PutValue(150);
            dataSheet.Cells["A5"].PutValue("Beverage");
            dataSheet.Cells["B5"].PutValue(70);
            dataSheet.Cells["A6"].PutValue("Supplies");
            dataSheet.Cells["B6"].PutValue(200);

            // -------------------------------------------------
            // Create a worksheet to host the pivot table and slicer
            // -------------------------------------------------
            Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");
            // Add a pivot table based on the data range
            int pivotIndex = pivotSheet.PivotTables.Add("A1:B6", "C3", "SalesPivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];
            // Configure pivot fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Sum of Amount

            // -------------------------------------------------
            // Add a slicer linked to the pivot table (Category field)
            // -------------------------------------------------
            int slicerIndex = pivotSheet.Slicers.Add(pivotTable, "A1", "Category");
            Slicer slicer = pivotSheet.Slicers[slicerIndex];
            slicer.Caption = "Category Filter";

            // Save initial state
            workbook.Save("SlicerVersion_0.xlsx");

            // -------------------------------------------------
            // First modification: change slicer caption
            // -------------------------------------------------
            slicer.Caption = "Product Category";
            slicer.Refresh(); // Refresh slicer and underlying pivot table
            workbook.Save("SlicerVersion_1.xlsx");

            // -------------------------------------------------
            // Second modification: lock slicer position
            // -------------------------------------------------
            slicer.LockedPosition = true;
            slicer.Refresh();
            workbook.Save("SlicerVersion_2.xlsx");

            // -------------------------------------------------
            // Third modification: change number of columns displayed
            // -------------------------------------------------
            slicer.NumberOfColumns = 3;
            slicer.Refresh();
            workbook.Save("SlicerVersion_3.xlsx");

            // -------------------------------------------------
            // Fourth modification: resize slicer
            // -------------------------------------------------
            slicer.WidthPixel = 250;
            slicer.HeightPixel = 180;
            slicer.Refresh();
            workbook.Save("SlicerVersion_4.xlsx");

            // -------------------------------------------------
            // Fifth modification: unlock slicer position
            // -------------------------------------------------
            slicer.LockedPosition = false;
            slicer.Refresh();
            workbook.Save("SlicerVersion_5.xlsx");

            // Cleanup
            workbook.Dispose();

            Console.WriteLine("Slicer modifications saved with incremental versioned files.");
        }
    }
}
