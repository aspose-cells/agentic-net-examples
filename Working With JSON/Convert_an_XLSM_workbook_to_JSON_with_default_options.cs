using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Paths for source XLSM file and destination JSON file
        string sourcePath = "input.xlsm";
        string destPath = "output.json";

        // Load the XLSM workbook
        Workbook workbook = new Workbook(sourcePath);

        // Save the workbook as JSON using default JsonSaveOptions
        JsonSaveOptions jsonOptions = new JsonSaveOptions();
        workbook.Save(destPath, jsonOptions);

        Console.WriteLine("Workbook successfully converted to JSON.");
    }
}