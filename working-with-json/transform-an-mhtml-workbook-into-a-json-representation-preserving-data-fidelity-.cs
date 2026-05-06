using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Json;

namespace MhtmlToJsonExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source MHTML workbook
            string mhtmlPath = "input.mht";

            Workbook workbook;

            if (File.Exists(mhtmlPath))
            {
                // Load the MHTML file into a Workbook instance
                workbook = new Workbook(mhtmlPath);
            }
            else
            {
                // Create a sample workbook if the MHTML file is not found
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "SampleSheet";
                workbook.Worksheets[0].Cells["A1"].PutValue("Sample Data");
            }

            // Configure JSON save options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                // Export as a JSON representation of the Excel file structure
                ToExcelStruct = true,
                // Ensure the output is always a JSON object even for a single sheet
                AlwaysExportAsJsonObject = true,
                // Preserve hierarchical relationships (e.g., tables, named ranges)
                ExportNestedStructure = true
            };

            // Save the workbook to a memory stream using the JSON options
            using (MemoryStream jsonStream = new MemoryStream())
            {
                workbook.Save(jsonStream, jsonOptions);
                jsonStream.Position = 0; // Reset stream position for reading

                // Read the JSON content from the stream
                string jsonResult = new StreamReader(jsonStream).ReadToEnd();

                // Output the JSON representation
                Console.WriteLine(jsonResult);
            }
        }
    }
}