using System;
using Aspose.Cells;

namespace AsposeCellsAlignmentDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Access cell A1 and set a sample value
            Cell cell = worksheet.Cells["A1"];
            cell.PutValue("Centered Text");

            // Get the current style of the cell
            Style style = cell.GetStyle();

            // Set horizontal alignment to Center
            style.HorizontalAlignment = TextAlignmentType.Center;

            // Set vertical alignment to Center (middle)
            style.VerticalAlignment = TextAlignmentType.Center;

            // Apply the modified style back to the cell
            cell.SetStyle(style);

            // Save the workbook to a file
            workbook.Save("AlignedCell.xlsx");
        }
    }
}