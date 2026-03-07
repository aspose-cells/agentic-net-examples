using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsJsonSmartMarkerDemo
{
    // Sample data class matching the JSON structure
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string City { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the Excel template that contains smart markers (e.g., &=$Data.Name)
            string templatePath = "Template.xlsx";

            // Load the template workbook (XLSX)
            Workbook workbook = new Workbook(templatePath);

            // Sample JSON string
            string json = "{\"Name\":\"John Doe\",\"Age\":30,\"City\":\"New York\"}";

            // Deserialize JSON into a .NET object (optional, demonstrates deserialization)
            Person person = JsonSerializer.Deserialize<Person>(json);

            // (Optional) Re‑serialize the object back to JSON if you need to modify it
            string jsonForSmartMarker = JsonSerializer.Serialize(person);

            // Initialize WorkbookDesigner and assign the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // Set the JSON data source for smart markers.
            // The first parameter is the name used in the smart marker (e.g., &=$Data.Name)
            designer.SetJsonDataSource("Data", jsonForSmartMarker);

            // Process all smart markers in the workbook
            designer.Process();

            // Save the populated workbook
            string outputPath = "Result.xlsx";
            workbook.Save(outputPath);
        }
    }
}