using System;
using Aspose.Cells;

namespace AsposeCellsNumberFormatDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a numeric value into cell A1
            Cell cell = sheet.Cells["A1"];
            cell.PutValue(1234.5678);

            // Retrieve the default style and display its Number property (built‑in format index)
            Style defaultStyle = cell.GetStyle();
            Console.WriteLine("Default Number format index: " + defaultStyle.Number);
            Console.WriteLine("Default formatted value: " + cell.StringValue);

            // Create a new style and set the Number property to 2 (decimal with two places: 0.00)
            Style newStyle = workbook.CreateStyle();
            newStyle.Number = 2; // Built‑in format index for "0.00"

            // Apply the new style to the cell
            cell.SetStyle(newStyle);

            // Verify the Number property after applying the style
            Style appliedStyle = cell.GetStyle();
            Console.WriteLine("Applied Number format index: " + appliedStyle.Number);
            Console.WriteLine("Formatted value after applying Number=2: " + cell.StringValue);

            // Save the workbook (lifecycle rule: use provided save method)
            workbook.Save("NumberFormatDemo.xlsx");
        }
    }
}