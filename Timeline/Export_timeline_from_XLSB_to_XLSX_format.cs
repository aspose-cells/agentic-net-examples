using System;
using System.IO;
using Aspose.Cells;

class ExportTimelineXlsbToXlsx
{
    static void Main()
    {
        // Path to the source XLSB file that contains the Timeline
        string sourceFile = Path.Combine(Directory.GetCurrentDirectory(), "input.xlsb");

        // Desired output path for the XLSX file
        string outputFile = Path.Combine(Directory.GetCurrentDirectory(), "output.xlsx");

        // If the source file does not exist, create a simple workbook and save it as XLSB
        if (!File.Exists(sourceFile))
        {
            var tempWb = new Workbook();
            tempWb.Worksheets[0].Name = "Sheet1";
            tempWb.Save(sourceFile, SaveFormat.Xlsb);
        }

        // Load the XLSB workbook (including any Timeline objects)
        var loadOptions = new LoadOptions(LoadFormat.Xlsb);
        var workbook = new Workbook(sourceFile, loadOptions);

        // Save the workbook as XLSX (OOXML)
        workbook.Save(outputFile, SaveFormat.Xlsx);

        Console.WriteLine($"Timeline exported from '{sourceFile}' to '{outputFile}'.");
    }
}