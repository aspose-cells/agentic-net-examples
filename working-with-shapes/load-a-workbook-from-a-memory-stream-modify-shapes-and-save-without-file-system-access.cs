using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsMemoryDemo
{
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Create a new workbook and add a rectangle shape.
            // ------------------------------------------------------------
            Workbook originalWorkbook = new Workbook();                     // create workbook
            Worksheet originalSheet = originalWorkbook.Worksheets[0];      // get first worksheet

            // Add a rectangle shape: (upper left row, column, upper left offset X, offset Y, width, height)
            Shape rect = originalSheet.Shapes.AddRectangle(1, 1, 0, 0, 100, 50);
            rect.Name = "OriginalRectangle";
            rect.FillFormat.ForeColor = Color.Red;                         // initial fill color

            // ------------------------------------------------------------
            // 2. Save the workbook to a memory stream (XLSX format).
            // ------------------------------------------------------------
            using (MemoryStream sourceStream = new MemoryStream())
            {
                originalWorkbook.Save(sourceStream, SaveFormat.Xlsx);      // save to stream
                sourceStream.Position = 0;                                 // reset for reading

                // ------------------------------------------------------------
                // 3. Load the workbook from the memory stream.
                // ------------------------------------------------------------
                Workbook loadedWorkbook = new Workbook(sourceStream);       // load from stream
                Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

                // ------------------------------------------------------------
                // 4. Modify the shape: rename and change fill color.
                // ------------------------------------------------------------
                if (loadedSheet.Shapes.Count > 0)
                {
                    Shape shapeToModify = loadedSheet.Shapes[0];
                    shapeToModify.Name = "ModifiedRectangle";
                    shapeToModify.FillFormat.ForeColor = Color.Blue;        // change fill color
                }

                // ------------------------------------------------------------
                // 5. Save the modified workbook to another memory stream.
                // ------------------------------------------------------------
                using (MemoryStream resultStream = new MemoryStream())
                {
                    loadedWorkbook.Save(resultStream, SaveFormat.Xlsx);   // save modified workbook
                    resultStream.Position = 0;                            // ready for further use

                    // Example usage: display the size of the resulting stream.
                    Console.WriteLine($"Modified workbook size: {resultStream.Length} bytes");
                }
            }
        }
    }
}