using System;
using Aspose.Cells;

namespace AsposeCellsDocumentPropertiesDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set built‑in document properties
            workbook.BuiltInDocumentProperties.Author = "John Smith";
            workbook.BuiltInDocumentProperties.Title = "Sample Workbook";
            workbook.BuiltInDocumentProperties.Company = "Aspose Ltd.";
            workbook.BuiltInDocumentProperties.CreatedTime = DateTime.Now;

            // Add custom document properties
            workbook.CustomDocumentProperties.Add("Project", "DocumentPropertiesDemo");
            workbook.CustomDocumentProperties.Add("Revision", 3);
            workbook.CustomDocumentProperties.Add("Approved", true);
            workbook.CustomDocumentProperties.Add("ReviewDate", DateTime.Today);

            // Display properties
            Console.WriteLine("Built‑in Properties:");
            Console.WriteLine($"Author: {workbook.BuiltInDocumentProperties.Author}");
            Console.WriteLine($"Title: {workbook.BuiltInDocumentProperties.Title}");
            Console.WriteLine($"Company: {workbook.BuiltInDocumentProperties.Company}");
            Console.WriteLine($"Created Time: {workbook.BuiltInDocumentProperties.CreatedTime}");

            Console.WriteLine("\nCustom Properties:");
            foreach (var prop in workbook.CustomDocumentProperties)
            {
                Console.WriteLine($"{prop.Name}: {prop.Value} ({prop.Type})");
            }

            // Save the workbook
            string outputPath = "DocumentPropertiesDemo.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            // Load the saved workbook to verify persisted properties
            Workbook loadedWorkbook = new Workbook(outputPath);

            Console.WriteLine("\nVerified after reload:");
            Console.WriteLine($"Author: {loadedWorkbook.BuiltInDocumentProperties.Author}");
            Console.WriteLine($"Custom Project: {loadedWorkbook.CustomDocumentProperties["Project"].Value}");
        }
    }
}