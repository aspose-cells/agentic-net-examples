using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class CustomPaperSizeHelper
    {
        /// <summary>
        /// Sets a custom paper size for the first worksheet.
        /// Width and height are provided in millimeters, converted to inches (the unit required by CustomPaperSize).
        /// </summary>
        /// <param name="widthMm">Paper width in millimeters.</param>
        /// <param name="heightMm">Paper height in millimeters.</param>
        public static void SetCustomPaperSize(double widthMm, double heightMm)
        {
            // Convert millimeters to inches (1 inch = 25.4 mm)
            double widthInches = widthMm / 25.4;
            double heightInches = heightMm / 25.4;

            // Create a new workbook (using the standard creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set the custom paper size (method expects inches)
            sheet.PageSetup.CustomPaperSize(widthInches, heightInches);

            // Save the workbook (using the standard save rule)
            workbook.Save("CustomPaperSize.xlsx");
        }

        // Example usage
        public static void Main()
        {
            // Example: set paper size to A5 (148 mm x 210 mm)
            SetCustomPaperSize(148, 210);
            Console.WriteLine("Custom paper size set and workbook saved.");
        }
    }
}