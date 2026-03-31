using System;
using System.IO;
using Aspose.Cells;

class PdfToJsonConverter
{
    static void Main()
    {
        string inputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "input.pdf");
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        var loadOptions = new LoadOptions(LoadFormat.Auto);
        var workbook = new Workbook(inputPath, loadOptions);

        workbook.CalculateFormula();

        var jsonOptions = new JsonSaveOptions
        {
            ToExcelStruct = true
        };

        string outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "output.json");
        workbook.Save(outputPath, jsonOptions);

        Console.WriteLine($"Conversion completed: {inputPath} -> {outputPath}");
    }
}