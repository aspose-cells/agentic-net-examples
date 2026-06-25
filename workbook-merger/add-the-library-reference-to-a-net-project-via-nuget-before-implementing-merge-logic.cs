// Add Aspose.Cells to the project via NuGet:
//   dotnet add package Aspose.Cells

using System;
using Aspose.Cells;

namespace MergeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample data before merging
            cells[0, 0].PutValue("Merged Header");
            cells[0, 1].PutValue("Column B");
            cells[0, 2].PutValue("Column C");
            cells[1, 0].PutValue("Row 2, Col A");
            cells[2, 0].PutValue("Row 3, Col A");

            // Merge cells A1:C1 (first row, first column, 1 row, 3 columns)
            cells.Merge(0, 0, 1, 3);

            // Save the workbook
            workbook.Save("MergedOutput.xlsx");
        }
    }
}