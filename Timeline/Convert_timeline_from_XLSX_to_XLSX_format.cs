using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsTimelineConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file that contains a Timeline control
            string sourcePath = "TimelineSource.xlsx";

            // Path where the converted XLSX file will be saved
            string destinationPath = "TimelineConverted.xlsx";

            // Ensure the source file exists; create a simple workbook if it does not.
            if (!File.Exists(sourcePath))
            {
                var wb = new Workbook();
                wb.Worksheets[0].Name = "Sheet1";
                wb.Save(sourcePath, SaveFormat.Xlsx);
            }

            // Convert the source workbook to the destination format (XLSX to XLSX)
            ConversionUtility.Convert(sourcePath, destinationPath);

            Console.WriteLine($"Timeline workbook successfully converted from '{sourcePath}' to '{destinationPath}'.");
        }
    }
}