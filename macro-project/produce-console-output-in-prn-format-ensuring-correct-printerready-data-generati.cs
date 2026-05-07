using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPrnDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Printer Ready (PRN) Demo");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // Configure PCL save options – PCL output can be saved with .prn extension
            PclSaveOptions pclOptions = new PclSaveOptions
            {
                // Each worksheet will be rendered on a single page
                OnePagePerSheet = true,
                // Start from the first page and render only one page (optional)
                PageIndex = 0,
                PageCount = 1
            };

            // Define the output PRN file name
            string prnFilePath = "output.prn";

            // Save the workbook as a PRN (PCL) file
            workbook.Save(prnFilePath, pclOptions);
            Console.WriteLine($"Workbook saved as PRN file: {prnFilePath}");

            // Read the generated PRN file bytes
            byte[] prnData = File.ReadAllBytes(prnFilePath);

            // Write the raw PRN data to the console (stdout)
            // This allows downstream processes that capture stdout to receive printer‑ready data
            using (Stream stdout = Console.OpenStandardOutput())
            {
                stdout.Write(prnData, 0, prnData.Length);
                stdout.Flush();
            }
        }
    }
}