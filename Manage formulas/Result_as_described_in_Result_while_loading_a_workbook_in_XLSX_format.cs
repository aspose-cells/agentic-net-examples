using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source XLSX file
        string inputPath = "input.xlsx";

        // Detect the file format (optional, demonstrates FileFormatUtil)
        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(inputPath);
        Console.WriteLine($"Detected LoadFormat: {formatInfo.LoadFormat}");

        // Create LoadOptions and configure it (e.g., skip formula parsing on open)
        LoadOptions loadOptions = new LoadOptions(formatInfo.LoadFormat);
        loadOptions.ParsingFormulaOnOpen = false;

        // Load the workbook using the constructor that accepts a file path and LoadOptions
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Console.WriteLine($"First worksheet name: {sheet.Name}");

        // Print the value of cell A1 (if any)
        Console.WriteLine($"Cell A1 value: {sheet.Cells["A1"].Value}");

        // Save the workbook to a new file
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to: {outputPath}");
    }
}