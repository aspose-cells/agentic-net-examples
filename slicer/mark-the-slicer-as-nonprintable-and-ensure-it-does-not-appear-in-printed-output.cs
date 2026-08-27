// Title: How to mark an Excel slicer as non‑printable with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to add a slicer to a pivot table and set its Shape.IsPrintable property to false so it won't appear in printed output. | Generate a complete Aspose.Cells example that creates sample data, builds a pivot table, inserts a linked slicer, and disables printing of that slicer before saving the workbook. | Provide a C# snippet that accesses the underlying Shape of a slicer in Aspose.Cells and configures IsPrintable = false, then saves the file as an .xlsx.
// Common Searches: Aspose.Cells C# set slicer non printable before saving workbook | How to prevent a slicer from printing in Excel using Aspose.Cells .NET | Example code for making slicer invisible in printed output with Aspose.Cells | C# Aspose.Cells slicer Shape.IsPrintable false usage | Exclude slicer from print area in generated Excel file Aspose.Cells
// Tags: Aspose.Cells slicer non printable | C# set slicer IsPrintable property | Excel pivot table slicer hide from print | Aspose.Cells Shape.IsPrintable usage | generate Excel file with non‑printable slicer

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

namespace AsposeCellsSlicerNonPrintableDemo
{
    // The example creates a workbook, adds sample data, builds a pivot table, inserts a slicer linked to the pivot, sets the slicer's underlying Shape.IsPrintable property to false to exclude it from printed output, and saves the workbook as SlicerNonPrintable.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].Value = "Category";
            cells["B1"].Value = "Amount";
            cells["A2"].Value = "A";
            cells["B2"].Value = 100;
            cells["A3"].Value = "B";
            cells["B3"].Value = 150;
            cells["A4"].Value = "C";
            cells["B4"].Value = 200;

            // Add a pivot table based on the data range
            PivotTableCollection pivots = sheet.PivotTables;
            int pivotIdx = pivots.Add("A1:B4", "D1", "Pivot1");
            PivotTable pivot = pivots[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer linked to the pivot table
            SlicerCollection slicers = sheet.Slicers;
            int slicerIdx = slicers.Add(pivot, "F1", "Category");
            Slicer slicer = slicers[slicerIdx];

            // Mark the slicer as non‑printable using the underlying Shape object
            slicer.Shape.IsPrintable = false;

            // Save the workbook
            workbook.Save("SlicerNonPrintable.xlsx");
        }
    }
}
