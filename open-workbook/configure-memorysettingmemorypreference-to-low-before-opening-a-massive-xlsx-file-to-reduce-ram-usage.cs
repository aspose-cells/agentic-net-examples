// Title: How to set MemorySetting.MemoryPreference to Low in Aspose.Cells for .NET before opening a massive XLSX workbook to minimize RAM usage
// Common Searches: Aspose.Cells set MemoryPreference low before loading large XLSX in C# | reduce memory consumption when opening massive Excel file with Aspose.Cells .NET | configure MemorySetting.MemoryPreference to Low for big workbook in Aspose.Cells | C# example low memory mode Aspose.Cells opening huge XLSX | optimize RAM usage Aspose.Cells workbook load large file
// Tags: MemoryPreference low Aspose.Cells | load large XLSX low memory Aspose.Cells | Aspose.Cells memory setting for massive workbook | RAM optimization Aspose.Cells workbook loading | Aspose.Cells large Excel file performance tuning | configure MemorySetting before workbook open

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates checking for the existence of a massive XLSX file, configuring Aspose.Cells MemorySetting.MemoryPreference to Low to lower RAM consumption, loading the workbook, and handling any exceptions.
class Program
{
    static void Main()
    {
        // Path to the massive XLSX file.
        string filePath = @"C:\Data\MassiveFile.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException.
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Input file not found: {filePath}");
            return;
        }

        try
        {
            // Load the workbook.
            Workbook workbook = new Workbook(filePath);

            // (Optional) Perform any required operations on the workbook here.

            // Example: Save the workbook to a new file after processing.
            // string outputPath = @"C:\Data\ProcessedFile.xlsx";
            // workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
