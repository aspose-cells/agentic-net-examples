using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source workbook. The file can be XLSX, XLS, CSV, etc.
        string sourcePath = "sample.xlsx";

        // The Workbook constructor that takes a file path automatically detects the format.
        Workbook workbook = new Workbook(sourcePath);

        // Display the detected file format.
        Console.WriteLine($"Detected file format: {workbook.FileFormat}");

        // Example processing: save the workbook to a different format (PDF).
        workbook.Save("converted.pdf", SaveFormat.Pdf);

        Console.WriteLine("Workbook processed and saved as PDF.");
    }
}