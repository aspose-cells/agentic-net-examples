using System;
using System.IO;
using Aspose.Cells;

class ConvertMhtToJson
{
    static void Main()
    {
        // Path to the source workbook (using XLSX for compatibility)
        string sourcePath = "input.xlsx";

        // If the workbook does not exist, create a sample workbook and save it
        if (!File.Exists(sourcePath))
        {
            Workbook sampleWb = new Workbook();
            Worksheet ws = sampleWb.Worksheets[0];
            ws.Cells["A1"].PutValue("Sample Text");
            ws.Cells["B1"].PutValue(12345);
            sampleWb.Save(sourcePath, SaveFormat.Xlsx);
        }

        // Path where the JSON output will be saved
        string outputPath = "output.json";

        // Load the workbook
        Workbook workbook = new Workbook(sourcePath);

        // Configure JSON save options
        JsonSaveOptions saveOptions = new JsonSaveOptions
        {
            // Export as a JSON object even if the workbook contains a single worksheet
            AlwaysExportAsJsonObject = true,
            // Export data using a nested (parent‑child) structure
            ExportNestedStructure = true
        };

        // Save the workbook as JSON
        workbook.Save(outputPath, saveOptions);

        Console.WriteLine("Workbook has been successfully converted to JSON.");
    }
}