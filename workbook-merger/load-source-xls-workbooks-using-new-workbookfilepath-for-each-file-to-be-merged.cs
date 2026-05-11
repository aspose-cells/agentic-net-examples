using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;

        string[] sourceFiles = new string[]
        {
            Path.Combine(baseDir, "Source1.xls"),
            Path.Combine(baseDir, "Source2.xls"),
            Path.Combine(baseDir, "Source3.xls")
        };

        Workbook destinationWorkbook = new Workbook();

        foreach (string filePath in sourceFiles)
        {
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                continue;
            }

            Workbook sourceWorkbook = new Workbook(filePath);
            destinationWorkbook.Combine(sourceWorkbook);
        }

        string outputPath = Path.Combine(baseDir, "MergedOutput.xls");
        destinationWorkbook.Save(outputPath);
        Console.WriteLine($"Merged workbook saved to: {outputPath}");
    }
}