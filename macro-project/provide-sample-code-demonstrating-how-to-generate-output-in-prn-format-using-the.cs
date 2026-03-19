using System;
using Aspose.Cells;

namespace AsposeCellsPrnDemo
{
    public class GeneratePrn
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Bananas");
            sheet.Cells["B3"].PutValue(85);

            // Optional: configure page setup (print area, fit to page, etc.)
            PageSetup pageSetup = sheet.PageSetup;
            pageSetup.PrintArea = "A1:B3";
            pageSetup.SetFitToPages(1, 1); // fit all data on one page

            // Save the workbook as a PRN (printer) file.
            // The format is inferred from the file extension.
            workbook.Save("SampleOutput.prn");

            Console.WriteLine("PRN file generated successfully: SampleOutput.prn");
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            GeneratePrn.Run();
        }
    }
}