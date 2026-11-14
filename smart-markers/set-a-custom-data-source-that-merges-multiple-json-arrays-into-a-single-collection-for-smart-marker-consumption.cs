// Title: How to merge multiple JSON arrays into a single data source for Aspose.Cells smart markers using C#
// Common Searches: c# merge two json arrays and bind to Aspose.Cells smart markers | set custom data source for WorkbookDesigner from combined json collections | using System.Text.Json to combine json arrays for smart markers in Aspose.Cells | example of WorkbookDesigner SetDataSource with merged json data in C# | aspnet cells smart marker data source from List<Dictionary> merged json
// Tags: json array merging for Aspose.Cells smart markers | WorkbookDesigner SetDataSource ListDictionary example | custom json collection data source C# | smart marker data binding using System.Text.Json | aspocells combined json data source

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// The sample creates a workbook, defines smart markers, merges two JSON arrays into a List of Dictionary objects using System.Text.Json, sets this merged collection as the "Data" source with WorkbookDesigner.SetDataSource, processes the smart markers, and saves the populated Excel file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add smart markers to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            // Smart markers that will be replaced by data from the custom source
            sheet.Cells["A2"].PutValue("&=Data.Name");
            sheet.Cells["B2"].PutValue("&=Data.Age");

            // Example JSON arrays (could be loaded from files, streams, etc.)
            string jsonArray1 = "[{\"Name\":\"John\",\"Age\":30},{\"Name\":\"Anna\",\"Age\":25}]";
            string jsonArray2 = "[{\"Name\":\"Mike\",\"Age\":40}]";

            // Merge the JSON arrays into a single collection suitable for smart markers
            List<Dictionary<string, object>> mergedData = new List<Dictionary<string, object>>();

            foreach (string json in new[] { jsonArray1, jsonArray2 })
            {
                using JsonDocument doc = JsonDocument.Parse(json);
                foreach (JsonElement element in doc.RootElement.EnumerateArray())
                {
                    var dict = new Dictionary<string, object>();
                    foreach (JsonProperty prop in element.EnumerateObject())
                    {
                        // Convert each JSON property to a .NET object
                        dict[prop.Name] = prop.Value.ValueKind == JsonValueKind.Null
                            ? null
                            : prop.Value.GetRawText();
                    }
                    mergedData.Add(dict);
                }
            }

            // Use WorkbookDesigner (compatible with all Aspose.Cells versions) to process smart markers
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            // The name "Data" matches the prefix used in the smart markers (&=Data.Name, etc.)
            designer.SetDataSource("Data", mergedData);
            designer.Process(); // Apply smart markers

            // Ensure the output directory exists
            string outputPath = "MergedJsonSmartMarkerOutput.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the resulting workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
