// Title: Duplicate an Excel worksheet with shapes while keeping the original Z‑order using Aspose.Cells for .NET
// Common Searches: Aspose.Cells copy worksheet with shapes preserving Z-order in C# | C# duplicate Excel sheet and retain shape stacking order using Aspose.Cells | How to maintain shape Z-order when copying a worksheet with Aspose.Cells .NET | Copy worksheet including charts and shapes without losing Z-order Aspose.Cells | Programmatically clone Excel worksheet and keep shape order in .NET
// Tags: Worksheet.Copy shape Z-order handling Aspose.Cells | duplicate sheet preserving shape hierarchy .NET | copy worksheet charts and shapes Aspose.Cells | shape stacking order preservation Aspose.Cells | clone Excel worksheet with shapes Aspose.Cells

using Aspose.Cells;
using System;

// The example loads a workbook, adds a new worksheet, and uses Worksheet.Copy to duplicate the source sheet—including cells, formats, charts, and all shapes—while preserving the original Z‑order of the shapes. An optional manual shape copy loop demonstrates how to enforce Z‑order explicitly, and the workbook is saved with the duplicated worksheet.
class Program
{
    static void Main()
    {
        // Load the source workbook
        Workbook workbook = new Workbook("source.xlsx");

        // Select the worksheet to duplicate (e.g., the first worksheet)
        Worksheet sourceSheet = workbook.Worksheets[0];

        // Add a new worksheet that will hold the duplicate
        Worksheet duplicateSheet = workbook.Worksheets.Add(sourceSheet.Name + "_Copy");

        // Duplicate the entire worksheet (cells, formats, charts, shapes, etc.)
        // Aspose.Cells copies shapes together with their Z‑order when using Worksheet.Copy
        sourceSheet.Copy(duplicateSheet);

        // If you need to ensure Z‑order explicitly, you can copy shapes manually
        // in the same order they appear in the source collection.
        // The following block demonstrates that approach (optional).

        //duplicateSheet.Shapes.Clear(); // Uncomment if you want to replace shapes copied by Copy()
        //foreach (Shape shape in sourceSheet.Shapes)
        //{
        //    // Add a copy of each shape; insertion order preserves Z‑order
        //    duplicateSheet.Shapes.AddCopy(shape);
        //}

        // Save the workbook with the duplicated worksheet
        workbook.Save("output.xlsx");
    }
}
