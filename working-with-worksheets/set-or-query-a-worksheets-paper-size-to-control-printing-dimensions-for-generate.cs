using System;
using Aspose.Cells;

namespace PaperSizeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default one worksheet)
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Set the default paper size for the whole workbook
            // This affects worksheets that do not have an explicit setting
            // -------------------------------------------------
            workbook.Settings.PaperSize = PaperSizeType.PaperA5;

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Set a specific paper size for this worksheet
            // This overrides the workbook default for this sheet only
            // -------------------------------------------------
            sheet.PageSetup.PaperSize = PaperSizeType.PaperLetter;

            // -------------------------------------------------
            // Query and display the paper size settings
            // -------------------------------------------------
            Console.WriteLine("Workbook default paper size: " + workbook.Settings.PaperSize);
            Console.WriteLine("Worksheet paper size: " + sheet.PageSetup.PaperSize);

            // -------------------------------------------------
            // Save the workbook to an XLSX file
            // -------------------------------------------------
            workbook.Save("PaperSizeDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}