using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsExamples
{
    class XmlToJsonConverter
    {
        static void Main()
        {
            // Path to the source XML Excel file
            string xmlPath = "source.xml";

            // Path for the resulting JSON file
            string jsonPath = "output.json";

            // If the source XML file does not exist, create a simple workbook and save it as XML
            if (!File.Exists(xmlPath))
            {
                Workbook tempWb = new Workbook();
                Worksheet sheet = tempWb.Worksheets[0];
                sheet.Name = "Sheet1";
                sheet.Cells["A1"].PutValue("Header");
                sheet.Cells["A2"].PutValue("Value1");
                sheet.Cells["B2"].PutValue("Value2");
                tempWb.Save(xmlPath, SaveFormat.Xml);
            }

            // Load the XML file using XmlLoadOptions (ensures proper handling of XML format)
            XmlLoadOptions loadOptions = new XmlLoadOptions();
            Workbook workbook = new Workbook(xmlPath, loadOptions);

            // Configure JSON save options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                // Export the workbook as a JSON object even if it contains a single worksheet
                AlwaysExportAsJsonObject = true,
                // Include header row if present
                HasHeaderRow = true,
                // Export empty cells as null (optional)
                ExportEmptyCells = true,
                // Indent the JSON for readability
                Indent = "  "
            };

            // Save the workbook as JSON using the configured options
            workbook.Save(jsonPath, jsonOptions);

            Console.WriteLine($"XML workbook '{xmlPath}' has been successfully converted to JSON at '{jsonPath}'.");
        }
    }
}