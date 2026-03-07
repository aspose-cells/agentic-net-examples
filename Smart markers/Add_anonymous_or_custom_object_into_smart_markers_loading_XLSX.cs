using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerDemo
{
    class Program
    {
        static void Main()
        {
            // Load the template workbook that contains smart markers (e.g., &Data.Name, &Data.Age)
            var workbook = new Workbook("template.xlsx");

            // Create a WorkbookDesigner and associate it with the loaded workbook
            var designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // Prepare a collection of anonymous objects as the data source.
            // Each object must have properties that match the smart marker field names.
            var customData = new[]
            {
                new
                {
                    Name = "John Doe",
                    Age = 30,
                    Photo = LoadPhoto("john_photo.jpg")
                },
                new
                {
                    Name = "Jane Smith",
                    Age = 28,
                    Photo = LoadPhoto("jane_photo.jpg")
                }
            };

            // Bind the anonymous object collection to a smart marker name (e.g., "Data")
            designer.SetDataSource("Data", customData);

            // Process all smart markers in the workbook
            designer.Process();

            // Save the populated workbook
            workbook.Save("output.xlsx");
        }

        private static byte[] LoadPhoto(string fileName)
        {
            return File.Exists(fileName) ? File.ReadAllBytes(fileName) : Array.Empty<byte>();
        }
    }
}