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
            // Step 1: Create a new workbook and set initial properties
            // -----------------------------------------------------------------
            string initialFile = "InitialWorkbook.xlsx";
            Workbook wb = new Workbook(); // create a new workbook
            // Set some built‑in properties directly on the workbook
            wb.BuiltInDocumentProperties.Author = "Original Author";
            wb.BuiltInDocumentProperties.Title = "Original Title";
            // Add a custom property
            wb.CustomDocumentProperties.Add("InitialCustom", "Value1");
            // Save the workbook to disk
            wb.Save(initialFile);

            // -----------------------------------------------------------------
            // Step 2: Load workbook metadata with DocumentProperties option
            // -----------------------------------------------------------------
            MetadataOptions options = new MetadataOptions(MetadataType.DocumentProperties);
            WorkbookMetadata metadata = new WorkbookMetadata(initialFile, options); // load metadata

            // -----------------------------------------------------------------
            // Step 3: Modify built‑in and custom properties via metadata API
            // -----------------------------------------------------------------
            // Built‑in properties are read‑write
            metadata.BuiltInDocumentProperties.Author = "Aspose Developer";
            metadata.BuiltInDocumentProperties.Title = "Metadata Updated Title";

            // Add or update custom properties
            metadata.CustomDocumentProperties.Add("UpdatedCustom", "NewValue");
            // (If the property already exists, you could also modify its Value)

            // -----------------------------------------------------------------
            // Step 4: Save the modified metadata back to the file
            // -----------------------------------------------------------------
            metadata.Save(initialFile); // overwrite the same file with updated metadata

            // -----------------------------------------------------------------
            // Step 5: Reload the workbook and verify that properties were saved
            // -----------------------------------------------------------------
            Workbook verifiedWb = new Workbook(initialFile);

            // Verify built‑in properties
            Console.WriteLine("Verified Author: " + verifiedWb.BuiltInDocumentProperties.Author);
            Console.WriteLine("Verified Title: " + verifiedWb.BuiltInDocumentProperties.Title);

            // Verify custom properties
            Console.WriteLine("Verified Custom Property 'InitialCustom': " +
                verifiedWb.CustomDocumentProperties["InitialCustom"].Value);
            Console.WriteLine("Verified Custom Property 'UpdatedCustom': " +
                verifiedWb.CustomDocumentProperties["UpdatedCustom"].Value);
        }
    }
}