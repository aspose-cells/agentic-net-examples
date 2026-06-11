using System;
using Aspose.Cells;

namespace AsposeCellsRowHeightExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the default row height to 20 points (standard height)
            worksheet.Cells.StandardHeight = 20;

            // Optional: add some sample data to visualize the height
            worksheet.Cells["A1"].PutValue("Row height set to 20 points");
            worksheet.Cells["A2"].PutValue("Another row with default height");

            // Save the workbook (lifecycle rule: save)
            workbook.Save("DefaultRowHeight20Points.xlsx");
        }
    }
}