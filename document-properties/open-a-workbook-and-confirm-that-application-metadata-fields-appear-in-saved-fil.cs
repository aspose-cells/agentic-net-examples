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
            // Step 1: Create a simple workbook and save it to disk
            Workbook workbook = new Workbook();                     // create a new workbook
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Metadata Demo");           // add some data
            string originalPath = "MetadataDemo_original.xlsx";
            workbook.Save(originalPath);                           // save the workbook (uses provided Save rule)

            // Step 2: Prepare metadata options to work with document properties
            MetadataOptions options = new MetadataOptions(MetadataType.DocumentProperties);

            // Step 3: Load the workbook metadata using the constructor (string, MetadataOptions)
            WorkbookMetadata metadata = new WorkbookMetadata(originalPath, options);

            // Step 4: Modify built‑in document properties
            BuiltInDocumentPropertyCollection builtInProps = metadata.BuiltInDocumentProperties;
            builtInProps.Author = "Aspose Developer";
            builtInProps.Title = "Metadata Demo Workbook";

            // Step 5: Add a custom document property
            metadata.CustomDocumentProperties.Add("Project", "MetadataAPI");

            // Step 6: Save the modified metadata back to a new file
            string updatedPath = "MetadataDemo_updated.xlsx";
            metadata.Save(updatedPath);                            // uses Save(string) rule

            // Step 7: Load the updated workbook to verify that properties were persisted
            Workbook verifiedWorkbook = new Workbook(updatedPath);

            // Verify built‑in properties
            Console.WriteLine("Verified Built‑in Properties:");
            Console.WriteLine("Author: " + verifiedWorkbook.BuiltInDocumentProperties.Author);
            Console.WriteLine("Title : " + verifiedWorkbook.BuiltInDocumentProperties.Title);

            // Verify custom property
            Console.WriteLine("\nVerified Custom Properties:");
            if (verifiedWorkbook.CustomDocumentProperties.Contains("Project"))
            {
                Console.WriteLine("Project: " + verifiedWorkbook.CustomDocumentProperties["Project"].Value);
            }
            else
            {
                Console.WriteLine("Custom property 'Project' not found.");
            }
        }
    }
}