using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsContentTypeDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing Excel file
            string filePath = "SampleWorkbook.xlsx";

            // Load the workbook using the constructor that accepts a file path
            Workbook workbook = new Workbook(filePath);

            // Access the ContentTypeProperties collection
            ContentTypePropertyCollection contentProps = workbook.ContentTypeProperties;

            // Check if any content type properties exist
            if (contentProps.Count == 0)
            {
                Console.WriteLine("No ContentTypeProperties found in the workbook.");
            }
            else
            {
                Console.WriteLine($"Found {contentProps.Count} ContentTypeProperty(ies):");
                // Iterate through each property and display its details
                foreach (ContentTypeProperty prop in contentProps)
                {
                    Console.WriteLine($"Name: {prop.Name}");
                    Console.WriteLine($"Value: {prop.Value}");
                    Console.WriteLine($"Type: {prop.Type}");
                    Console.WriteLine($"IsNillable: {prop.IsNillable}");
                    Console.WriteLine(new string('-', 40));
                }
            }

            // Dispose the workbook when done
            workbook.Dispose();
        }
    }
}