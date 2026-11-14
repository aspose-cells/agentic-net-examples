// Title: Save an Excel workbook as ODS and verify file size reduction with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells save workbook to ODS and compare file size with original XLSX | How to export an XLSX file to ODS using Aspose.Cells and check if the ODS is smaller | Aspose.Cells .NET code to convert Excel to ODS and output size difference | Determine size reduction when converting Excel workbook to ODS with Aspose.Cells in C# | Measure file size before and after saving workbook as ODS using Aspose.Cells
// Tags: Aspose.Cells ODS export C# | XLSX to ODS size comparison Aspose.Cells | convert Excel workbook to ODS .NET | verify file size reduction after ODS conversion | range manipulation before ODS save Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads an existing XLSX file into an Aspose.Cells Workbook, optionally manipulates a cell range, saves the workbook as an ODS file, then reads the file sizes of the original and converted files and prints whether the ODS file is smaller.
class Program
{
    static void Main()
    {
        // Path to the original Excel file
        string inputPath = "input.xlsx";

        // Load the workbook (create/load rule applied)
        Workbook workbook = new Workbook(inputPath);

        // Optionally, you could manipulate a specific range here
        // Example: Worksheet sheet = workbook.Worksheets[0];
        // Range range = sheet.Cells.CreateRange("A1:C10");
        // ... (conversion logic if needed)

        // Path for the ODS output file
        string odsPath = "output.ods";

        // Save the workbook as ODS (save rule applied)
        workbook.Save(odsPath, SaveFormat.Ods);

        // Get file sizes to confirm reduction
        long originalSize = new FileInfo(inputPath).Length;
        long odsSize = new FileInfo(odsPath).Length;

        // Output the size comparison
        Console.WriteLine($"Original XLSX size: {originalSize} bytes");
        Console.WriteLine($"Converted ODS size: {odsSize} bytes");

        if (odsSize < originalSize)
        {
            Console.WriteLine("File size reduction confirmed.");
        }
        else
        {
            Console.WriteLine("No size reduction detected.");
        }
    }
}
