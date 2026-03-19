using System;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells.Charts;

namespace AsposeCellsJsonDemo
{
    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // 1. Create a workbook and populate sample data
            // -------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet ws = workbook.Worksheets[0];

            ws.Cells["A1"].PutValue("Name");
            ws.Cells["B1"].PutValue("Age");
            ws.Cells["A2"].PutValue("John");
            ws.Cells["B2"].PutValue(30);
            ws.Cells["A3"].PutValue("Alice");
            ws.Cells["B3"].PutValue(25);

            // -------------------------------------------------
            // 2. Export a specific range to JSON using JsonSaveOptions
            // -------------------------------------------------
            Aspose.Cells.Range exportRange = ws.Cells.CreateRange("A1:B3");
            JsonSaveOptions exportOptions = new JsonSaveOptions
            {
                Indent = "    ",               // 4‑space indentation
                HasHeaderRow = true,          // First row contains column names
                ExportEmptyCells = true,      // Include empty cells in output
                ExportStylePool = false       // Export styles per cell
            };
            string jsonFromRange = JsonUtility.ExportRangeToJson(exportRange, exportOptions);
            Console.WriteLine("Exported JSON from range:");
            Console.WriteLine(jsonFromRange);

            // -------------------------------------------------
            // 3. Save the workbook as a JSON file using the same options
            // -------------------------------------------------
            string jsonFilePath = "ExportedData.json";
            workbook.Save(jsonFilePath, exportOptions);
            Console.WriteLine($"Workbook saved as JSON to '{jsonFilePath}'.");

            // -------------------------------------------------
            // 4. Import JSON data into a new workbook using JsonLayoutOptions
            // -------------------------------------------------
            string jsonToImport = @"{
                ""Products"": [
                    { ""ID"": 101, ""Name"": ""Product A"", ""Price"": 99.99 },
                    { ""ID"": 102, ""Name"": ""Product B"", ""Price"": 149.50 }
                ]
            }";

            Workbook importWorkbook = new Workbook();
            Worksheet importWs = importWorkbook.Worksheets[0];
            JsonLayoutOptions layoutOptions = new JsonLayoutOptions
            {
                ArrayAsTable = true,
                KeptSchema = true               // Preserve schema if needed later
            };
            JsonUtility.ImportData(jsonToImport, importWs.Cells, 0, 0, layoutOptions);
            Console.WriteLine("JSON data imported into new workbook.");

            // -------------------------------------------------
            // 5. Demonstrate smart markers with JSON data source
            // -------------------------------------------------
            // Add a marker cell that will be replaced by the JSON data
            importWs.Cells["D1"].PutValue("&=$Products.Name");
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = importWorkbook
            };
            // Use the same JSON string as data source for the marker
            designer.SetJsonDataSource("Products", jsonToImport);
            designer.Process();
            Console.WriteLine("Smart markers processed with JSON data source.");

            // -------------------------------------------------
            // 6. Save the imported workbook to JSON with schema validation
            // -------------------------------------------------
            string schema = @"{
                ""$schema"": ""http://json-schema.org/draft-07/schema#"",
                ""type"": ""object"",
                ""properties"": {
                    ""Products"": {
                        ""type"": ""array"",
                        ""items"": {
                            ""type"": ""object"",
                            ""properties"": {
                                ""ID"": { ""type"": ""integer"" },
                                ""Name"": { ""type"": ""string"" },
                                ""Price"": { ""type"": ""number"" }
                            },
                            ""required"": [""ID"", ""Name"", ""Price""]
                        }
                    }
                },
                ""required"": [""Products""]
            }";

            JsonSaveOptions saveWithSchema = new JsonSaveOptions
            {
                Schemas = new[] { schema },
                ExportNestedStructure = true,
                Indent = "\t",                 // Tab indentation for readability
                ExportStylePool = true
            };
            string finalJsonPath = "FinalOutput.json";
            importWorkbook.Save(finalJsonPath, saveWithSchema);
            Console.WriteLine($"Final workbook saved as JSON with schema to '{finalJsonPath}'.");

            // -------------------------------------------------
            // 7. Optional: Create a chart and set values format code
            // -------------------------------------------------
            int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = ws.Charts[chartIdx];
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";
            chart.NSeries[0].ValuesFormatCode = "0.00%"; // Display values as percentages
            Console.WriteLine("Chart created and values format code set.");

            // -------------------------------------------------
            // End of demonstration
            // -------------------------------------------------
        }
    }
}