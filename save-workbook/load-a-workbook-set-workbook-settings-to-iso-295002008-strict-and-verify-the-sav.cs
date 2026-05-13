using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the workbook that will be loaded
        string sourcePath = "source.xlsx";

        // Load the workbook from the file system (uses Workbook(string) constructor)
        Workbook workbook = new Workbook(sourcePath);

        // Access the workbook settings
        WorkbookSettings settings = workbook.Settings;

        // Set the OOXML compliance level to ISO/IEC 29500:2008 Strict
        settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

        // Define the path for the saved workbook
        string outputPath = "output_strict.xlsx";

        // Save the workbook using the Save(string) method
        workbook.Save(outputPath);

        // Verify the saved file size
        FileInfo fileInfo = new FileInfo(outputPath);
        Console.WriteLine($"Saved file size: {fileInfo.Length} bytes");
    }
}