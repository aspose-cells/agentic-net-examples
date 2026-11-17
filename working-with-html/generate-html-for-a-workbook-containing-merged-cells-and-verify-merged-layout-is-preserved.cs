// Title: C# example: merge cells A1:C2, save to MemoryStream, reload workbook, and confirm merged layout with Aspose.Cells
// Common Searches: Aspose.Cells C# merge cells A1:C2 and verify after reloading workbook | how to check merged cell areas after saving to a MemoryStream using Aspose.Cells | load workbook from stream and inspect merged cell ranges with Aspose.Cells for .NET | sample code to ensure merged layout is preserved when saving as Xlsx in Aspose.Cells | C# verify merged cells collection after workbook reload Aspose.Cells
// Tags: merge cells Aspose.Cells C# | verify merged cell area Aspose.Cells | save workbook to MemoryStream Aspose.Cells | load workbook from stream Aspose.Cells | iterate CellArea collection Aspose.Cells | preserve merged layout Xlsx Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace MergedCellsDemo
{
    // Demonstrates creating a workbook, merging cells A1:C2, saving it to a MemoryStream in Xlsx format, reloading the workbook, and confirming that the merged cell area is retained by iterating the CellArea collection.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Merge a range of cells (A1:C2)
            // Parameters: first row, first column, total rows, total columns
            sheet.Cells.Merge(0, 0, 2, 3); // Merges cells A1:C2

            // Save the workbook to a memory stream (could be a file as well)
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, SaveFormat.Xlsx);
                ms.Position = 0; // Reset stream position for reading

                // Load the workbook from the memory stream
                Workbook loadedWorkbook = new Workbook(ms);
                Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

                // Verify that the merged cells layout is preserved
                bool mergePreserved = false;

                // Iterate through all merged cell areas
                foreach (CellArea area in loadedSheet.Cells.MergedCells)
                {
                    // Check if the expected merged area (A1:C2) exists
                    if (area.StartRow == 0 && area.StartColumn == 0 &&
                        area.EndRow == 1 && area.EndColumn == 2)
                    {
                        mergePreserved = true;
                        break;
                    }
                }

                // Output verification result
                Console.WriteLine("Merged cells layout preserved: " + mergePreserved);
            }
        }
    }
}
