using System;
using Aspose.Cells;

namespace AsposeCellsSxcExample
{
    public class Program
    {
        public static void Main()
        {
            // Path to the source SXC file
            string sourcePath = "input.sxc";

            // Path for the modified SXC file
            string destPath = "output.sxc";

            // Load the workbook from the SXC file
            Workbook workbook = new Workbook(sourcePath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Modify existing cells or add new data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1.25);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(0.75);

            // Save the workbook back to SXC format
            workbook.Save(destPath, SaveFormat.Sxc);
        }
    }
}