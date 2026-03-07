using System;
using Aspose.Cells;

class LoadXlsxExample
{
    static void Main()
    {
        // Path to the source XLSX file
        string inputPath = "input.xlsx";

        // Create LoadOptions specifying the Xlsx format explicitly
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

        // Load the workbook using the file path and the load options
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Display basic information about the loaded workbook
        Console.WriteLine($"Loaded file: {inputPath}");
        Console.WriteLine($"Number of worksheets: {workbook.Worksheets.Count}");

        // Save the workbook to a new file to verify successful load and save
        string outputPath = "output_copy.xlsx";
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved as: {outputPath}");
    }
}