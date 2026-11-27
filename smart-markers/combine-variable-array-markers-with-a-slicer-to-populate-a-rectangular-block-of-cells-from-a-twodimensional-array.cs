// Title: Use a variable array smart marker slicer to fill a rectangular block from a 2‑D array in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# variable array smart marker slicer example to populate B2:C4 | How to bind a two‑dimensional object array to a smart marker in Aspose.Cells | Insert smart marker '&Items[0:2,0:1]' to copy a sub‑range from a 2D array in Excel using Aspose.Cells | Populate specific rows and columns from a 2D array with Aspose.Cells WorkbookDesigner | C# code for slicing a 2D array into an Excel range using Aspose.Cells smart markers
// Tags: smart marker array slicer syntax Aspose.Cells | populate Excel cells from 2D object array C# | WorkbookDesigner bind 2D array Aspose.Cells | C# Aspose.Cells subrange smart marker | Excel range filling with variable array marker

using System;
using Aspose.Cells;

namespace SmartMarkerArraySlicerDemo
{
    // The example creates a 4×4 string[,] array, places a variable array smart marker with slicer '&Items[0:2,0:1]' in cell B2, binds the array to the marker name 'Items' via WorkbookDesigner, processes the markers, and saves the workbook, resulting in the 3×2 block B2:C4 being filled.
    class Program
    {
        static void Main(string[] args)
        {
            // ------------------------------------------------------------
            // 1. Prepare a two‑dimensional array that will be used as a data source.
            // ------------------------------------------------------------
            // The array can contain any type that Aspose.Cells can write to a cell
            // (string, double, DateTime, etc.). Here we use strings for clarity.
            object[,] sourceArray = new object[,]
            {
                { "R1C1", "R1C2", "R1C3", "R1C4" },
                { "R2C1", "R2C2", "R2C3", "R2C4" },
                { "R3C1", "R3C2", "R3C3", "R3C4" },
                { "R4C1", "R4C2", "R4C3", "R4C4" }
            };

            // ------------------------------------------------------------
            // 2. Create a new workbook and obtain the first worksheet.
            // ------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // 3. Insert a variable array marker with a slicer.
            // ------------------------------------------------------------
            // The syntax "&DataName[StartRow:EndRow,StartCol:EndCol]" tells the
            // designer to copy a rectangular block from the source array.
            // In this example we want to copy rows 1‑3 (index 0‑2) and columns
            // 1‑2 (index 0‑1), i.e. a 3x2 block.
            // The marker is placed in the top‑left cell where the block will start.
            sheet.Cells["B2"].PutValue("&Items[0:2,0:1]");

            // ------------------------------------------------------------
            // 4. Create a WorkbookDesigner, bind the data source and process the markers.
            // ------------------------------------------------------------
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            // Bind the two‑dimensional array to the marker name "Items".
            designer.SetDataSource("Items", sourceArray);
            // Process all smart markers in the workbook.
            designer.Process();

            // ------------------------------------------------------------
            // 5. Save the result so you can verify the populated block.
            // ------------------------------------------------------------
            string outputPath = "ArraySlicerResult.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to '{outputPath}'.");
            Console.WriteLine("The block populated from the 2‑D array occupies cells B2:C4.");
        }
    }
}
