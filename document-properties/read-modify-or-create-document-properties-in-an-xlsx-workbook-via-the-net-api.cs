using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsDocumentPropertiesDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (or create a new one if the file does not exist)
            // Lifecycle rule: use the Workbook constructor for loading.
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // ------------------------------
            // Modify built‑in document properties
            // Example: change Author and Title
            // BuiltInDocumentProperties is a collection; properties are accessed by name.
            // Lifecycle rule: use the existing collection, do not create a new one.
            workbook.BuiltInDocumentProperties["Author"].Value = "Jane Doe";
            workbook.BuiltInDocumentProperties["Title"].Value = "Quarterly Report";

            // ------------------------------
            // Add or update a custom document property
            // First, check if the property already exists.
            const string customPropName = "ReviewedBy";
            if (workbook.CustomDocumentProperties.Contains(customPropName))
            {
                // Update existing property value
                workbook.CustomDocumentProperties[customPropName].Value = "John Smith";
            }
            else
            {
                // Add a new custom property (string type)
                // Lifecycle rule: use the Add method of CustomDocumentPropertyCollection.
                workbook.CustomDocumentProperties.Add(customPropName, "John Smith");
            }

            // Add another custom property of a different type (DateTime)
            workbook.CustomDocumentProperties.Add("ReviewDate", DateTime.Now);

            // ------------------------------
            // Save the workbook with the updated properties
            // Lifecycle rule: use the Save method.
            workbook.Save("OutputWorkbook.xlsx", SaveFormat.Xlsx);

            // Optional: display the updated properties to the console
            Console.WriteLine("Updated Built‑in Properties:");
            Console.WriteLine($"Author: {workbook.BuiltInDocumentProperties["Author"].Value}");
            Console.WriteLine($"Title: {workbook.BuiltInDocumentProperties["Title"].Value}");

            Console.WriteLine("\nCustom Properties:");
            foreach (DocumentProperty prop in workbook.CustomDocumentProperties)
            {
                Console.WriteLine($"{prop.Name}: {prop.Value} ({prop.Type})");
            }
        }
    }
}