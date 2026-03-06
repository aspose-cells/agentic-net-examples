using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Directory containing the Excel file (current directory)
        string dataDir = Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar;

        // Input XLSX file to load
        string inputPath = Path.Combine(dataDir, "sample.xlsx");

        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        // Load the workbook from the XLSX file
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Display worksheet information
        Console.WriteLine("Worksheet Name: " + worksheet.Name);
        Console.WriteLine("Cell A1 Value: " + worksheet.Cells["A1"].StringValue);

        // Save the workbook to a new XLSX file
        string outputPath = Path.Combine(dataDir, "output.xlsx");
        workbook.Save(outputPath, SaveFormat.Xlsx);

        Console.WriteLine($"Workbook saved to: {outputPath}");
    }
}