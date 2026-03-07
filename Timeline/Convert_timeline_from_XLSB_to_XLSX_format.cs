using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source XLSB file
        string sourcePath = "TimelineSource.xlsb";

        // Desired output path for the converted XLSX file
        string destinationPath = "TimelineConverted.xlsx";

        // Ensure the source file exists; if not, create a simple workbook and save as XLSB
        if (!File.Exists(sourcePath))
        {
            var wbCreate = new Workbook();
            wbCreate.Worksheets[0].Name = "Sheet1";
            wbCreate.Save(sourcePath, SaveFormat.Xlsb);
        }

        // Load the XLSB workbook and save it as XLSX
        var workbook = new Workbook(sourcePath);
        workbook.Save(destinationPath, SaveFormat.Xlsx);

        Console.WriteLine($"Conversion completed: '{sourcePath}' -> '{destinationPath}'");
    }
}