using System;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Path to the source XLSX file
        string sourcePath = "input.xlsx";

        // Desired path for the resulting JSON file
        string destPath = "output.json";

        // Convert the Excel workbook to JSON using default conversion settings
        ConversionUtility.Convert(sourcePath, destPath);

        Console.WriteLine($"Conversion completed: '{sourcePath}' -> '{destPath}'.");
    }
}