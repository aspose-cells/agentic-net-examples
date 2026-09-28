// Title: Create versioned Excel files after each slicer property change using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that adds a slicer to a PivotTable, updates its caption, lock state, column count, and dimensions, refreshes the slicer, and saves the workbook to sequentially numbered .xlsx files. | Show how to implement incremental versioning of an Excel workbook by invoking Workbook.Save with a dynamic filename after each slicer adjustment in Aspose.Cells.
// Common Searches: Aspose.Cells C# save workbook after each slicer update with versioned filenames | how to implement incremental Excel file versioning when changing slicer properties in .NET | C# example for tracking slicer changes by saving separate Excel files using Aspose.Cells | automate slicer property modifications and file versioning in Aspose.Cells for .NET | generate sequentially named Excel files after slicer adjustments with Aspose.Cells
// Tags: Aspose.Cells save workbook after slicer change | incremental Excel versioning Aspose.Cells C# | slicer property modification Aspose.Cells | refresh slicer linked pivot Aspose.Cells | dynamic filename generation C# Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerVersioningDemo
{
    // The sample program creates a workbook with sample data, builds a PivotTable, adds a slicer linked to the PivotTable, then sequentially modifies slicer properties (caption, locked position, column count, size). After each modification the slicer is refreshed and the workbook is saved to a uniquely numbered file (SlicerVersion_0.xlsx through SlicerVersion_5.xlsx), demonstrating incremental versioning for change tracking.
    public class Program
    {
        public static void Main()
        {
            // Initialize a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // 1. Prepare sample data for a PivotTable
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
            dataSheet.Cells["B5"].PutValue(200);
            dataSheet.Cells["A6"].PutValue("Supplies");
            dataSheet.Cells["B6"].PutValue(90);

            // -------------------------------------------------
            // 2. Create a PivotTable based on the data
            // -------------------------------------------------
            Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");
            int pivotIndex = pivotSheet.PivotTables.Add("A1:B6", "C3", "SalesPivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Sum of Amount

            // -------------------------------------------------
            // 3. Add a slicer linked to the PivotTable
            // -------------------------------------------------
            Worksheet slicerSheet = workbook.Worksheets.Add("Slicer");
            int slicerIndex = slicerSheet.Slicers.Add(pivotTable, "A1", "Category");
            Slicer slicer = slicerSheet.Slicers[slicerIndex];

            // Initial save (version 0)
            workbook.Save("SlicerVersion_0.xlsx");

            // -------------------------------------------------
            // 4. Modify slicer properties and save after each change
            // -------------------------------------------------

            // Change 1: Update caption
            slicer.Caption = "Category Filter";
            slicer.Refresh(); // Refresh slicer and underlying PivotTable
            workbook.Save("SlicerVersion_1.xlsx");

            // Change 2: Lock the slicer position
            slicer.LockedPosition = true;
            slicer.Refresh();
            workbook.Save("SlicerVersion_2.xlsx");

            // Change 3: Set number of columns displayed in the slicer
            slicer.NumberOfColumns = 2;
            slicer.Refresh();
            workbook.Save("SlicerVersion_3.xlsx");

            // Change 4: Adjust size of the slicer
            slicer.WidthPixel = 250;
            slicer.HeightPixel = 180;
            slicer.Refresh();
            workbook.Save("SlicerVersion_4.xlsx");

            // Change 5: Unlock the slicer position
            slicer.LockedPosition = false;
            slicer.Refresh();
            workbook.Save("SlicerVersion_5.xlsx");

            // Cleanup
            workbook.Dispose();

            Console.WriteLine("Slicer modifications saved with incremental versioned files.");
        }
    }
}
