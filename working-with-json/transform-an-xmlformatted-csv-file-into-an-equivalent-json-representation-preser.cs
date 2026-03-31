using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class XmlCsvToJson
{
    static void Main()
    {
        // Path to the XML‑formatted CSV file (replace with your actual file path)
        string xmlCsvPath = "data.xml";

        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Import the XML data into the first worksheet starting at cell A1 (row 0, column 0)
        workbook.ImportXml(xmlCsvPath, "Sheet1", 0, 0);

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Determine the used range that contains the imported data
        Aspose.Cells.Range usedRange = worksheet.Cells.MaxDisplayRange;

        // Configure JSON export options to preserve data types and hierarchical structure
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            ExportEmptyCells = true,      // keep empty cells in the output
            HasHeaderRow = true,          // treat the first row as header
            ExportNestedStructure = true, // preserve nested structures if any
            ToExcelStruct = true          // output in Excel‑like JSON struct
        };

        // Export the used range to a JSON string using the configured options
        string jsonOutput = JsonUtility.ExportRangeToJson(usedRange, jsonOptions);

        // Write the JSON string to a file
        File.WriteAllText("output.json", jsonOutput);

        Console.WriteLine("XML‑formatted CSV has been converted to JSON and saved as output.json");
    }
}