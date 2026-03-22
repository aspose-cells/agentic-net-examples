using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source XLSB file containing the timeline worksheet
        string sourcePath = "timeline.xlsb";

        // Desired path for the converted XLSX file
        string destinationPath = "timeline.xlsx";

        // Ensure the source file exists; create a dummy file if it does not.
        if (!File.Exists(sourcePath))
        {
            var dummyWorkbook = new Workbook();
            dummyWorkbook.Worksheets[0].Name = "Timeline";
            dummyWorkbook.Save(sourcePath, SaveFormat.Xlsb);
        }

        // Load the XLSB workbook.
        var loadOptions = new LoadOptions(LoadFormat.Xlsb);
        var workbook = new Workbook(sourcePath, loadOptions);

        // Save the workbook as XLSX.
        workbook.Save(destinationPath, SaveFormat.Xlsx);

        Console.WriteLine($"Successfully converted '{sourcePath}' to '{destinationPath}'.");
    }
}