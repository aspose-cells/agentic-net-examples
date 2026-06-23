using System;
using Aspose.Cells;

namespace FitToPagesWideLogger
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the PageSetup object
            PageSetup pageSetup = worksheet.PageSetup;

            // Log the original FitToPagesWide value
            int originalValue = pageSetup.FitToPagesWide;
            Console.WriteLine($"Original FitToPagesWide: {originalValue}");

            // Change the FitToPagesWide value
            // Example: set to 2 pages wide
            pageSetup.FitToPagesWide = 2;

            // Log the updated FitToPagesWide value
            int updatedValue = pageSetup.FitToPagesWide;
            Console.WriteLine($"Updated FitToPagesWide: {updatedValue}");

            // Optionally, use SetFitToPages method to demonstrate another way of setting
            // This will also affect FitToPagesWide (set to 3) and FitToPagesTall (set to 0)
            pageSetup.SetFitToPages(3, 0);
            Console.WriteLine($"After SetFitToPages - FitToPagesWide: {pageSetup.FitToPagesWide}, FitToPagesTall: {pageSetup.FitToPagesTall}");

            // Save the workbook (lifecycle rule: save)
            workbook.Save("FitToPagesWideLog.xlsx");
        }
    }
}