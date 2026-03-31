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
            // -----------------------------------------------------------------
            // 1. Create a simple workbook and save it to a file.
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();                     // create a new workbook
            workbook.Worksheets[0].Cells["A1"].PutValue("Demo");   // add some data
            string filePath = "Sample.xlsx";
            workbook.Save(filePath);                               // save the workbook

            // -----------------------------------------------------------------
            // 2. Load the workbook metadata for document properties.
            // -----------------------------------------------------------------
            MetadataOptions options = new MetadataOptions(MetadataType.DocumentProperties);
            WorkbookMetadata metadata = new WorkbookMetadata(filePath, options); // load metadata

            // -----------------------------------------------------------------
            // 3. Modify built‑in document properties (read‑write).
            // -----------------------------------------------------------------
            BuiltInDocumentPropertyCollection builtInProps = metadata.BuiltInDocumentProperties;
            builtInProps.Author = "Aspose Developer";
            builtInProps.Title = "Metadata Demo Workbook";
            builtInProps.Subject = "Demonstrating metadata API";
            builtInProps.Company = "Aspose Ltd.";
            builtInProps.Comments = "Updated built‑in properties via WorkbookMetadata";

            // -----------------------------------------------------------------
            // 4. Add or modify custom document properties.
            // -----------------------------------------------------------------
            CustomDocumentPropertyCollection customProps = metadata.CustomDocumentProperties;
            // Add a new custom property
            customProps.Add("Project", "MetadataDemo");
            // Add another custom property with a numeric value
            customProps.Add("Revision", 2);
            // Update an existing custom property (if it already exists)
            if (customProps.Contains("Project"))
            {
                customProps["Project"].Value = "MetadataDemoUpdated";
            }

            // -----------------------------------------------------------------
            // 5. Save the modified metadata back to the workbook file.
            // -----------------------------------------------------------------
            metadata.Save(filePath); // overwrites the original file with updated metadata

            // -----------------------------------------------------------------
            // 6. Verify the changes by re‑loading the workbook and reading properties.
            // -----------------------------------------------------------------
            Workbook verifiedWorkbook = new Workbook(filePath);

            // Built‑in properties are accessible directly from the Workbook instance
            Console.WriteLine("Verified Built‑In Properties:");
            Console.WriteLine($"Author:  {verifiedWorkbook.BuiltInDocumentProperties["Author"].Value}");
            Console.WriteLine($"Title:   {verifiedWorkbook.BuiltInDocumentProperties["Title"].Value}");
            Console.WriteLine($"Subject: {verifiedWorkbook.BuiltInDocumentProperties["Subject"].Value}");
            Console.WriteLine($"Company: {verifiedWorkbook.BuiltInDocumentProperties["Company"].Value}");
            Console.WriteLine($"Comments:{verifiedWorkbook.BuiltInDocumentProperties["Comments"].Value}");

            // Custom properties
            Console.WriteLine("\nVerified Custom Properties:");
            foreach (DocumentProperty prop in verifiedWorkbook.CustomDocumentProperties)
            {
                Console.WriteLine($"{prop.Name}: {prop.Value} (Type: {prop.Type})");
            }
        }
    }
}