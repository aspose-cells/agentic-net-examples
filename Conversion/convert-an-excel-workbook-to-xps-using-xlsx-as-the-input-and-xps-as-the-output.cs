using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Desired path for the output XPS file
            string destPath = "output.xps";

            // Perform the conversion using Aspose.Cells ConversionUtility
            ConversionUtility.Convert(sourcePath, destPath);

            Console.WriteLine("Conversion from XLSX to XPS completed successfully.");
        }
    }
}