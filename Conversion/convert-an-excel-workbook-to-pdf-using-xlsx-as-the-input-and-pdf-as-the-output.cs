using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        string sourcePath = Path.Combine(Environment.CurrentDirectory, "input.xlsx");
        string destPath = Path.Combine(Environment.CurrentDirectory, "output.pdf");

        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Source file not found: {sourcePath}");
            return;
        }

        Workbook workbook = new Workbook(sourcePath);
        workbook.Save(destPath, SaveFormat.Pdf);

        Console.WriteLine($"Conversion completed successfully. PDF saved to {destPath}");
    }
}