using System;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file that may contain Office Add‑Ins
            string sourcePath = "input.xlsx";

            // Desired output PDF file path
            string destPath = "output.pdf";

            // Convert the Excel file to PDF using Aspose.Cells ConversionUtility.
            // This method handles all necessary loading and saving internally,
            // and no additional conversion options are required.
            ConversionUtility.Convert(sourcePath, destPath);

            Console.WriteLine("Conversion completed successfully.");
        }
    }
}