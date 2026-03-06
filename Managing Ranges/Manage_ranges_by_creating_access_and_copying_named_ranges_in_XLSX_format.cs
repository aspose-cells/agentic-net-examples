using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeDemo
{
    class Program
    {
        static void Main()
        {
            // Create workbook and define a named range
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // Fill sample data in A1:C3
            for (int row = 0; row < 3; row++)
                for (int col = 0; col < 3; col++)
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");

            // Create a range object for A1:C3 and assign a name to the range object
            AsposeRange srcRange = cells.CreateRange("A1", "C3");
            srcRange.Name = "SourceData";

            // Also create a workbook‑level named range that refers to the same cells
            int nameIdx = wb.Worksheets.Names.Add("MyNamedRange");
            Name namedRange = wb.Worksheets.Names[nameIdx];
            namedRange.RefersTo = "=Sheet1!$A$1:$C$3";

            // Save the workbook
            wb.Save("RangeDemo.xlsx");

            // Load workbook, access the named range and copy it
            Workbook loadedWb = new Workbook("RangeDemo.xlsx");
            Worksheet loadedWs = loadedWb.Worksheets[0];
            Cells loadedCells = loadedWs.Cells;

            // Retrieve the range via the Name object (GetRange method)
            Name loadedNamedRange = loadedWb.Worksheets.Names["MyNamedRange"];
            AsposeRange accessedRange = loadedNamedRange.GetRange();

            // Define a destination range starting at E1 with the same size (3 rows x 3 columns)
            AsposeRange destRange = loadedCells.CreateRange(0, 4, 3, 3); // row 0, column 4 => cell E1

            // Copy the source range to the destination range
            destRange.Copy(accessedRange);

            // Change the name of the copied range
            destRange.Name = "CopiedData";

            // Save the modified workbook
            loadedWb.Save("RangeDemo_Copied.xlsx");
        }
    }
}