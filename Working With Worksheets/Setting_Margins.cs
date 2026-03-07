using System;
using Aspose.Cells;

namespace AsposeCellsMarginDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data to visualize margins
            sheet.Cells["A1"].PutValue("Margin Demonstration");
            sheet.Cells["A2"].PutValue("Top, Bottom, Left, Right margins are set.");

            // Access the PageSetup object
            PageSetup pageSetup = sheet.PageSetup;

            // Set margins in centimeters
            pageSetup.TopMargin = 2.0;      // 2 cm top margin
            pageSetup.BottomMargin = 1.5;   // 1.5 cm bottom margin
            pageSetup.LeftMargin = 1.0;     // 1 cm left margin
            pageSetup.RightMargin = 1.0;    // 1 cm right margin

            // Alternatively, set margins in inches using the Inch properties
            // pageSetup.TopMarginInch = 0.79; // approx 2 cm
            // pageSetup.BottomMarginInch = 0.59; // approx 1.5 cm
            // pageSetup.LeftMarginInch = 0.39; // approx 1 cm
            // pageSetup.RightMarginInch = 0.39; // approx 1 cm

            // Save the workbook to PDF to see the effect of margins
            workbook.Save("MarginDemo.pdf", SaveFormat.Pdf);

            Console.WriteLine("Workbook saved with custom margins.");
        }
    }
}