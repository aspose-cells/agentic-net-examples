using System;
using System.IO;
using System.Collections;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the template file into a byte array (replace with your source as needed)
        byte[] templateBytes = File.ReadAllBytes("template.xlsx");

        // Create a memory stream from the template bytes
        using (MemoryStream templateStream = new MemoryStream(templateBytes))
        {
            // Load the workbook from the stream
            Workbook workbook = new Workbook(templateStream);

            // Initialize WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // Prepare a sample data source for smart markers
            var persons = new ArrayList
            {
                new Person { Name = "John Doe", Age = 30 },
                new Person { Name = "Jane Smith", Age = 28 }
            };

            // Set the data source (the name must match the smart marker prefix in the template)
            designer.SetDataSource("Persons", persons);

            // Process the smart markers and populate the workbook
            designer.Process();

            // Save the processed workbook to a memory stream (returns a MemoryStream)
            MemoryStream resultStream = workbook.SaveToStream();

            // Convert the memory stream to a byte array
            byte[] resultBytes = resultStream.ToArray();

            // (Optional) Write the result to a file for verification
            File.WriteAllBytes("result.xlsx", resultBytes);
        }
    }

    // Simple POCO class used as a data source for smart markers
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }
}