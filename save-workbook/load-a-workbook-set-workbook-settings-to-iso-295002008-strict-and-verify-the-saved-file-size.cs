using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source workbook that will be loaded.
        // Replace with an actual file path as needed.
        string sourcePath = "source.xlsx";

        // Load the existing workbook using the constructor that accepts a file name.
        Workbook workbook = new Workbook(sourcePath);

        // Set the OOXML compliance level to ISO/IEC 29500:2008 Strict.
        workbook.Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

        // Define the output file path.
        string outputPath = "output_strict.xlsx";

        // Save the workbook using the standard Save(string) method.
        workbook.Save(outputPath);

        // Verify the saved file size.
        long fileSize = new FileInfo(outputPath).Length;
        Console.WriteLine($"Saved file size: {fileSize} bytes");
    }
}