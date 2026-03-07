using System;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Path to the source XLSB file (which may contain JSON data)
        string sourcePath = "input.xlsb";

        // Desired output CSV file path
        string destinationPath = "output.csv";

        // Convert the XLSB file to CSV using Aspose.Cells ConversionUtility
        // This overload automatically determines the formats from file extensions.
        ConversionUtility.Convert(sourcePath, destinationPath);

        Console.WriteLine("Conversion from XLSB to CSV completed successfully.");
    }
}