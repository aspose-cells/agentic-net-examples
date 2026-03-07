using System;
using Aspose.Cells;
using Aspose.Cells.Json; // JsonSaveOptions resides in this namespace

class Program
{
    static void Main()
    {
        // Path to the source XLS file
        string sourcePath = "input.xls";

        // Desired path for the JSON output
        string outputPath = "output.json";

        // Load the workbook from the XLS file
        Workbook workbook = new Workbook(sourcePath);

        // Save the workbook as JSON using default JsonSaveOptions
        workbook.Save(outputPath, new JsonSaveOptions());

        Console.WriteLine("Workbook successfully converted to JSON.");
    }
}