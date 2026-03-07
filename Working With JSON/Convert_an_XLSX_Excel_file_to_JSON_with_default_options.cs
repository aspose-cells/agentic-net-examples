using System;
using Aspose.Cells;

public class Program
{
    public static void Main()
    {
        Run();
    }

    public static void Run()
    {
        string sourcePath = "input.xlsx";
        string destPath = "output.json";

        Workbook workbook = new Workbook(sourcePath);
        JsonSaveOptions saveOptions = new JsonSaveOptions();
        workbook.Save(destPath, saveOptions);

        Console.WriteLine($"Conversion completed: {sourcePath} -> {destPath}");
    }
}