using System;
using System.IO;
using Aspose.Cells;

class PreserveFormulasJsonDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate cells with values and a formula
            worksheet.Cells["A1"].PutValue(10);
            worksheet.Cells["A2"].PutValue(20);
            worksheet.Cells["A3"].Formula = "=SUM(A1:A2)";

            // Configure JSON save options (ExportFormulas not available in this version)
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportEmptyCells = true,    // Include empty cells in the output
                HasHeaderRow = false        // Do not generate a header row
            };

            // Define JSON output path
            string jsonFilePath = "WorkbookWithFormulas.json";

            // Ensure the target directory exists
            string jsonDir = Path.GetDirectoryName(jsonFilePath);
            if (!string.IsNullOrEmpty(jsonDir) && !Directory.Exists(jsonDir))
                Directory.CreateDirectory(jsonDir);

            // Save the workbook as JSON
            workbook.Save(jsonFilePath, jsonOptions);

            // Convert a specific range to a JSON string using the same options
            Aspose.Cells.Range range = worksheet.Cells.CreateRange("A1:A3");
            string rangeJson = range.ToJson(jsonOptions);

            // Output the JSON string
            Console.WriteLine("Range JSON with formulas preserved:");
            Console.WriteLine(rangeJson);
        }
        catch (FileNotFoundException fnfEx)
        {
            Console.WriteLine($"File not found: {fnfEx.FileName}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}