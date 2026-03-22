using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        string sourcePath = "timeline_source.xlsm";
        string destinationPath = "timeline_converted.xlsx";

        if (!File.Exists(sourcePath))
        {
            var wb = new Workbook();
            wb.Worksheets[0].Name = "Sheet1";
            wb.Settings.EnableMacros = true;
            wb.Save(sourcePath, SaveFormat.Xlsm);
        }

        ConversionUtility.Convert(sourcePath, destinationPath);

        var convertedWb = new Workbook(destinationPath);
        Console.WriteLine($"Conversion completed. Workbook contains {convertedWb.Worksheets.Count} worksheet(s).");
    }
}