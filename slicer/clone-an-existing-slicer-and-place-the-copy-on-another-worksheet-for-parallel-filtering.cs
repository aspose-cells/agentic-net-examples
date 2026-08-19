// Title: Clone a Slicer to a Different Worksheet with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook with a pivot table, add a slicer on the source sheet, then duplicate that slicer on another worksheet while preserving caption, style, column layout, size, and locked‑position settings, and finally save the file.
// Keywords: Aspose.Cells slicer clone | duplicate slicer C# | copy slicer properties .NET | pivot table slicer replication | Aspose.Cells dashboard slicer | C# Excel slicer automation
// Common Searches: how to copy a slicer to another sheet using Aspose.Cells | C# duplicate slicer visual settings | Aspose.Cells clone slicer example | programmatically copy slicer style in .NET | create linked slicer on multiple worksheets
// Developer Intent: The developer needs to programmatically replicate an existing slicer on a different worksheet while keeping its visual configuration intact.
// Use Cases: Add identical slicers to several dashboard tabs so users can filter the same pivot table from any view. | Generate report templates where a master slicer is reused on detail sheets without manual re‑formatting. | Maintain consistent slicer appearance across worksheets in automated Excel generation workflows.
// AI Prompts: Write C# code with Aspose.Cells that clones a slicer to a specified worksheet and copies all visual properties. | Explain how to keep slicer selections synchronized across multiple sheets in an Aspose.Cells workbook. | Create a reusable method that takes an existing slicer and returns a cloned instance on another worksheet, preserving style, size, and lock settings.

using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// Demonstrates how to create a workbook with a pivot table, add a slicer on the source sheet, then duplicate that slicer on another worksheet while preserving caption, style, column layout, size, and locked‑position settings, and finally save the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sourceSheet = workbook.Worksheets[0];
        sourceSheet.Name = "Source";

        // Populate sample data for the pivot table
        sourceSheet.Cells["A1"].PutValue("Fruit");
        sourceSheet.Cells["B1"].PutValue("Sales");
        sourceSheet.Cells["A2"].PutValue("Apple");
        sourceSheet.Cells["B2"].PutValue(100);
        sourceSheet.Cells["A3"].PutValue("Orange");
        sourceSheet.Cells["B3"].PutValue(150);
        sourceSheet.Cells["A4"].PutValue("Banana");
        sourceSheet.Cells["B4"].PutValue(200);

        // Add a pivot table based on the data
        int pivotIdx = sourceSheet.PivotTables.Add("A1:B4", "D2", "Pivot1");
        PivotTable pivot = sourceSheet.PivotTables[pivotIdx];
        pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
        pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Add an original slicer on the source worksheet
        int slicerIdx = sourceSheet.Slicers.Add(pivot, "F2", "Fruit");
        Slicer originalSlicer = sourceSheet.Slicers[slicerIdx];
        originalSlicer.Caption = "Fruit Filter";
        originalSlicer.StyleType = SlicerStyleType.SlicerStyleLight2;
        originalSlicer.NumberOfColumns = 1;
        originalSlicer.WidthPixel = 150;
        originalSlicer.HeightPixel = 200;
        originalSlicer.LockedPosition = false;

        // Add a second worksheet where the cloned slicer will be placed
        Worksheet targetSheet = workbook.Worksheets.Add("Clone");

        // Create a new slicer on the target worksheet using the same pivot table and field
        int clonedIdx = targetSheet.Slicers.Add(pivot, "A1", "Fruit");
        Slicer clonedSlicer = targetSheet.Slicers[clonedIdx];

        // Copy visual properties from the original slicer to the cloned slicer
        clonedSlicer.Caption = originalSlicer.Caption;
        clonedSlicer.StyleType = originalSlicer.StyleType;
        clonedSlicer.NumberOfColumns = originalSlicer.NumberOfColumns;
        clonedSlicer.WidthPixel = originalSlicer.WidthPixel;
        clonedSlicer.HeightPixel = originalSlicer.HeightPixel;
        clonedSlicer.LockedPosition = originalSlicer.LockedPosition;

        // Save the workbook
        workbook.Save("ClonedSlicerDemo.xlsx");
    }
}
