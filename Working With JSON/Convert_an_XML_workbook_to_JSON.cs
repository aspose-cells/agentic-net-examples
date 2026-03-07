using System;
using Aspose.Cells;

class XmlToJsonConverter
{
    static void Main()
    {
        // Path to the source XML file that contains the data to be imported
        string xmlPath = "input.xml";

        // Name of the worksheet where the XML data will be placed
        string sheetName = "Sheet1";

        // Create a new, empty workbook
        Workbook workbook = new Workbook();

        // Import the XML data into the workbook starting at cell A1 (row 0, column 0)
        workbook.ImportXml(xmlPath, sheetName, 0, 0);

        // Configure JSON export options
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            // Export the workbook as a JSON object even if there is only one worksheet
            AlwaysExportAsJsonObject = true,
            // Preserve hierarchical (parent‑child) structure if present
            ExportNestedStructure = true,
            // Treat the first row as header names
            HasHeaderRow = true
        };

        // Path for the resulting JSON file
        string jsonPath = "output.json";

        // Save the workbook as JSON using the configured options
        workbook.Save(jsonPath, jsonOptions);
    }
}