using System;
using Aspose.Cells;

namespace AsposeCellsMHtmlConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Path for the output MHTML file
            string destPath = "output.mht";

            // Load the workbook from the XLSX file
            Workbook workbook = new Workbook(sourcePath);

            // Save the workbook as MHTML (MHT) preserving formatting
            workbook.Save(destPath, SaveFormat.MHtml);

            Console.WriteLine("Conversion to MHTML completed successfully.");
        }
    }
}