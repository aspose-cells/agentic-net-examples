using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Metadata;
using Aspose.Cells.Properties;

namespace AdvancedWorkbookPropertiesDemo
{
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a new workbook and set built‑in & custom document properties
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook(); // create workbook (uses provided ctor)

            // Set built‑in properties
            workbook.BuiltInDocumentProperties["Author"].Value = "Alice Johnson";
            workbook.BuiltInDocumentProperties["Title"].Value = "Quarterly Sales Report";
            workbook.BuiltInDocumentProperties["Company"].Value = "Acme Corp";

            // Add custom properties
            workbook.CustomDocumentProperties.Add("ReportId", 20230915);
            workbook.CustomDocumentProperties.Add("Reviewed", true);
            workbook.CustomDocumentProperties.Add("Department", "Finance");

            // Save the workbook (uses provided Save method)
            string filePath = "AdvancedPropertiesDemo.xlsx";
            workbook.Save(filePath);
            workbook.Dispose();

            // -----------------------------------------------------------------
            // 2. Load the workbook metadata to manipulate properties at the package level
            // -----------------------------------------------------------------
            // Create MetadataOptions to work with document properties
            MetadataOptions metaOptions = new MetadataOptions(MetadataType.DocumentProperties);

            // Load metadata from the existing file (uses provided ctor)
            WorkbookMetadata metadata = new WorkbookMetadata(filePath, metaOptions);

            // Modify built‑in properties via metadata (read‑write)
            metadata.BuiltInDocumentProperties.Author = "Bob Smith";
            metadata.BuiltInDocumentProperties.Title = "Updated Quarterly Sales Report";
            metadata.BuiltInDocumentProperties.Comments = "Reviewed and approved.";

            // Add additional custom property via metadata
            metadata.CustomDocumentProperties.Add("ApprovalDate", DateTime.Now);

            // Save the modified metadata back to the same file (uses provided Save method)
            metadata.Save(filePath);

            // -----------------------------------------------------------------
            // 3. Verify the changes by re‑loading the workbook
            // -----------------------------------------------------------------
            Workbook verifyWorkbook = new Workbook(filePath);
            Console.WriteLine("=== Verified Built‑In Properties ===");
            Console.WriteLine("Author: " + verifyWorkbook.BuiltInDocumentProperties["Author"].Value);
            Console.WriteLine("Title: " + verifyWorkbook.BuiltInDocumentProperties["Title"].Value);
            Console.WriteLine("Comments: " + verifyWorkbook.BuiltInDocumentProperties["Comments"].Value);
            Console.WriteLine("Company: " + verifyWorkbook.BuiltInDocumentProperties["Company"].Value);

            Console.WriteLine("\n=== Verified Custom Properties ===");
            foreach (DocumentProperty prop in verifyWorkbook.CustomDocumentProperties)
            {
                Console.WriteLine($"{prop.Name}: {prop.Value} ({prop.Type})");
            }

            verifyWorkbook.Dispose();
        }
    }
}