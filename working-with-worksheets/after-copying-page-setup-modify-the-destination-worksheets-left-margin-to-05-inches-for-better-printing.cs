using System;
using Aspose.Cells;

namespace AsposeCellsPageSetupExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook with two worksheets
            Workbook workbook = new Workbook();
            workbook.Worksheets.Add(); // Adds a second worksheet (index 1)

            // Access source and destination worksheets
            Worksheet sourceSheet = workbook.Worksheets[0];
            Worksheet destSheet = workbook.Worksheets[1];

            // Optionally set a custom left margin on the source sheet (in inches)
            sourceSheet.PageSetup.LeftMarginInch = 1.0; // 1 inch

            // Copy all page setup settings from source to destination
            destSheet.PageSetup.Copy(sourceSheet.PageSetup, new CopyOptions());

            // After copying, adjust the left margin of the destination sheet to 0.5 inches
            destSheet.PageSetup.LeftMarginInch = 0.5; // 0.5 inch for better printing

            // Save the workbook to verify the changes
            workbook.Save("PageSetupCopyAndMargin.xlsx");
        }
    }
}