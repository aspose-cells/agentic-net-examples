using System;
using Aspose.Cells;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Define source Excel file (XLSX) and destination XPS file paths
            string sourcePath = "input.xlsx";
            string destPath = "output.xps";

            // Load the workbook from the XLSX file (uses workbook-load rule)
            Workbook workbook = new Workbook(sourcePath);

            // Save the workbook as XPS (uses workbook-save rule with SaveFormat.Xps)
            workbook.Save(destPath, SaveFormat.Xps);

            Console.WriteLine("Conversion from XLSX to XPS completed successfully.");
        }
    }
}