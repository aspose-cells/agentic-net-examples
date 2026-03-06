using System;
using Aspose.Cells;
using Aspose.Cells.Metadata;

namespace AdvancedDocumentPropertiesDemo
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook
            Workbook workbook = new Workbook();

            // 2. Set built‑in document properties
            workbook.BuiltInDocumentProperties.Author = "Alice Johnson";
            workbook.BuiltInDocumentProperties.Title = "Quarterly Sales Report";
            workbook.BuiltInDocumentProperties.Company = "Contoso Ltd.";
            workbook.BuiltInDocumentProperties.CreatedTime = DateTime.Now;

            // 3. Add custom document properties
            workbook.CustomDocumentProperties.Add("ReviewedBy", "Bob Smith");
            workbook.CustomDocumentProperties.Add("ReviewDate", DateTime.Today);
            workbook.CustomDocumentProperties.Add("Version", 1.2);
            workbook.CustomDocumentProperties.Add("Approved", true);

            // 4. Add Content‑Type properties (metadata stored in the package)
            int ctIndex1 = workbook.ContentTypeProperties.Add("ProjectCode", "PRJ-2024");
            workbook.ContentTypeProperties[ctIndex1].IsNillable = true;

            int ctIndex2 = workbook.ContentTypeProperties.Add("ReportGenerated", DateTime.UtcNow.ToString("o"));
            workbook.ContentTypeProperties[ctIndex2].IsNillable = false;

            // 5. Save the workbook
            string originalPath = "AdvancedProperties.xlsx";
            workbook.Save(originalPath);
            workbook.Dispose();

            // 6. Load the workbook metadata
            MetadataOptions metaOptions = new MetadataOptions(MetadataType.DocumentProperties);
            WorkbookMetadata metadata = new WorkbookMetadata(originalPath, metaOptions);

            // 7. Read some built‑in properties via metadata
            var builtInMeta = metadata.BuiltInDocumentProperties;
            Console.WriteLine("Metadata - Author: " + builtInMeta.Author);
            Console.WriteLine("Metadata - Title: " + builtInMeta.Title);

            // 8. Modify built‑in properties
            builtInMeta.Author = "Aspose Developer";
            builtInMeta.Title = "Updated Quarterly Sales Report";

            // 9. Add a new custom property via metadata
            metadata.CustomDocumentProperties.Add("Department", "Finance");

            // 10. Save the modified metadata to a new file
            string updatedPath = "AdvancedProperties_Updated.xlsx";
            metadata.Save(updatedPath);
            metadata = null; // Release metadata object

            // 11. Verify the changes by loading the updated workbook
            Workbook updatedWorkbook = new Workbook(updatedPath);
            Console.WriteLine("Verified Author: " + updatedWorkbook.BuiltInDocumentProperties.Author);
            Console.WriteLine("Verified Title: " + updatedWorkbook.BuiltInDocumentProperties.Title);
            Console.WriteLine("Verified Custom Property 'Department': " +
                updatedWorkbook.CustomDocumentProperties["Department"].Value);
            updatedWorkbook.Dispose();
        }
    }
}