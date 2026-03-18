using System;
using System.IO;
using Aspose.Cells;

class ExcelToJsonConverter
{
    static void Main()
    {
        // Load the existing Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure JSON save options to export the whole workbook structure
        JsonSaveOptions saveOptions = new JsonSaveOptions
        {
            // Export as a JSON object even if there is only one worksheet
            AlwaysExportAsJsonObject = true,
            // Preserve the hierarchical (parent‑child) structure of sheets, rows and cells
            ExportNestedStructure = true,
            // Convert the Excel file into its JSON struct representation
            ToExcelStruct = true,
            // Optional: omit completely empty rows from the output
            SkipEmptyRows = true
        };

        // Save the workbook as a JSON file using the configured options
        string jsonPath = "output.json";
        workbook.Save(jsonPath, saveOptions);

        // Read and display the generated JSON (for verification)
        string jsonContent = File.ReadAllText(jsonPath);
        Console.WriteLine(jsonContent);
    }
}