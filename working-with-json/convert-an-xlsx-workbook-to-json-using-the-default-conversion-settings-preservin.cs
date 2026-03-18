using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source XLSX file
        string sourcePath = "input.xlsx";

        // Path where the JSON output will be saved
        string jsonPath = "output.json";

        // Load the workbook from the XLSX file
        using (Workbook workbook = new Workbook(sourcePath))
        {
            // Create default JSON save options (preserves worksheet structure and cell types)
            JsonSaveOptions saveOptions = new JsonSaveOptions();

            // Save the workbook as JSON using the default options
            workbook.Save(jsonPath, saveOptions);
        }

        Console.WriteLine("Workbook has been converted to JSON successfully.");
    }
}