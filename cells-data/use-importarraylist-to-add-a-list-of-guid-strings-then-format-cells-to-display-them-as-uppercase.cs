using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsGuidImport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet's cells collection
            Workbook workbook = new Workbook();
            Cells cells = workbook.Worksheets[0].Cells;

            // Prepare an ArrayList with GUID strings
            ArrayList guidList = new ArrayList();
            for (int i = 0; i < 10; i++)
            {
                guidList.Add(Guid.NewGuid().ToString()); // e.g., "d3c9a1e2-5b4f-4c9a-8f2e-1a2b3c4d5e6f"
            }

            // Import the GUID list vertically starting at cell A1 (row 0, column 0)
            // isVertical = true means each GUID will occupy a new row in the same column
            cells.ImportArrayList(guidList, 0, 0, true);

            // Convert all imported GUID strings to uppercase
            // The imported range occupies guidList.Count rows in column 0
            for (int row = 0; row < guidList.Count; row++)
            {
                // Retrieve the cell, ensure it contains a string, then set it to uppercase
                Cell cell = cells[row, 0];
                if (cell.Type == CellValueType.IsString)
                {
                    cell.PutValue(cell.StringValue.ToUpper());
                }
            }

            // Save the workbook to a file
            workbook.Save("GuidListUppercase.xlsx");
        }
    }
}