using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        string sourcePath = "input.sxc";
        string destPath = "output.csv";

        if (!System.IO.File.Exists(sourcePath))
        {
            Console.WriteLine($"Source file not found: {sourcePath}");
            return;
        }

        var loadOptions = new LoadOptions(LoadFormat.SXC);
        var workbook = new Workbook(sourcePath, loadOptions);
        workbook.Save(destPath, SaveFormat.CSV);

        Console.WriteLine($"Conversion completed: {sourcePath} -> {destPath}");
    }
}