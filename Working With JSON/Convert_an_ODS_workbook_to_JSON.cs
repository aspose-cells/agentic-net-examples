using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class OdsToJsonConverter
{
    static void Main()
    {
        // Path to the source ODS file
        string odsPath = "input.ods";

        // Path where the resulting JSON will be saved
        string jsonPath = "output.json";

        // Load the ODS workbook using OdsLoadOptions (default options)
        OdsLoadOptions loadOptions = new OdsLoadOptions();
        Workbook workbook = new Workbook(odsPath, loadOptions);

        // Configure JSON save options
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            // Export the workbook as a JSON object even if it contains a single worksheet
            AlwaysExportAsJsonObject = true,

            // Preserve the Excel structure (worksheets, tables, etc.) in the JSON output
            ToExcelStruct = true
        };

        // Save the workbook as JSON using the configured options
        workbook.Save(jsonPath, jsonOptions);

        Console.WriteLine("ODS workbook has been successfully converted to JSON.");
    }
}