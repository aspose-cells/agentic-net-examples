using System;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Path to the source Excel file (XLS)
        string sourcePath = "input.xls";

        // Desired path for the resulting JSON file
        string outputPath = "output.json";

        // Convert the Excel file to JSON using default conversion options
        ConversionUtility.Convert(sourcePath, outputPath);

        Console.WriteLine("Excel file has been successfully converted to JSON.");
    }
}