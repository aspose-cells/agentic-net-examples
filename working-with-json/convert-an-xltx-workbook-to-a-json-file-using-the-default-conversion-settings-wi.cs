using System;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Path to the source XLTX workbook
        string sourcePath = "template.xltx";

        // Desired output JSON file path
        string destPath = "output.json";

        // Convert the XLTX workbook to JSON using default conversion settings
        ConversionUtility.Convert(sourcePath, destPath);

        Console.WriteLine("Conversion completed successfully.");
    }
}