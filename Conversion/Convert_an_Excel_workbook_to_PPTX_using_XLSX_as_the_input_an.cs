using System;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file (XLSX)
            string sourcePath = "input.xlsx";

            // Desired output path with PPTX extension
            string destPath = "output.pptx";

            // Convert the Excel workbook to PowerPoint presentation
            // This uses Aspose.Cells.Utility.ConversionUtility.Convert(string, string)
            ConversionUtility.Convert(sourcePath, destPath);

            Console.WriteLine("Conversion completed successfully.");
        }
    }
}