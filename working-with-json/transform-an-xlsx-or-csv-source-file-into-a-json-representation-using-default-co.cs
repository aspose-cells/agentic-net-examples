using System;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Path to the source Excel or CSV file
        string sourceFile = "input.xlsx"; // change to "input.csv" if needed

        // Path where the JSON output will be saved
        string jsonFile = "output.json";

        // Convert the source file to JSON using default conversion settings
        ConversionUtility.Convert(sourceFile, jsonFile);

        Console.WriteLine($"Conversion completed: {sourceFile} -> {jsonFile}");
    }
}