using System;
using Aspose.Cells;
using Aspose.Cells.Metadata;
using Aspose.Cells.Properties;

namespace AsposeCellsMetadataDemo
{
    class Program
    {
        static void Main()
        {
            // Path for the workbook file
            string filePath = "MetadataDemo.xlsx";

            // -------------------------------------------------
            // 1. Create a new workbook and set initial properties
            // -------------------------------------------------
            Workbook workbook = new Workbook();

            // Set built‑in document properties
            workbook.BuiltInDocumentProperties.Author = "Initial Author";
            workbook.BuiltInDocumentProperties.Title = "Initial Title";

            // Add custom document properties
            workbook.CustomDocumentProperties.Add("Project", "MetadataDemo");
            workbook.CustomDocumentProperties.Add("Version", 1);

            // Save the workbook to disk
            workbook.Save(filePath);

            // -------------------------------------------------
            // 2. Load workbook metadata for document properties
            // -------------------------------------------------
            MetadataOptions options = new MetadataOptions(MetadataType.DocumentProperties);
            WorkbookMetadata metadata = new WorkbookMetadata(filePath, options);

            // Modify built‑in properties via metadata
            metadata.BuiltInDocumentProperties.Author = "Updated Author";
            metadata.BuiltInDocumentProperties.Title = "Updated Title";

            // Add/Update custom properties via metadata
            metadata.CustomDocumentProperties.Add("ReviewedBy", "John Doe");
            // Update existing custom property
            if (metadata.CustomDocumentProperties.Contains("Version"))
            {
                metadata.CustomDocumentProperties["Version"].Value = 2;
            }

            // Save the modified metadata back to the same file
            metadata.Save(filePath);

            // -------------------------------------------------
            // 3. Reload the workbook and verify properties
            // -------------------------------------------------
            Workbook loadedWorkbook = new Workbook(filePath);

            // Verify built‑in properties
            Console.WriteLine("Built‑in Author: " + loadedWorkbook.BuiltInDocumentProperties.Author);
            Console.WriteLine("Built‑in Title: " + loadedWorkbook.BuiltInDocumentProperties.Title);

            // Verify custom properties
            foreach (DocumentProperty prop in loadedWorkbook.CustomDocumentProperties)
            {
                Console.WriteLine($"Custom Property - Name: {prop.Name}, Value: {prop.Value}");
            }
        }
    }
}