// Title: Persist Built‑in and Custom Document Properties Using Aspose.Cells WorkbookMetadata (C#)
// Description: Demonstrates how to create an Excel workbook, set built‑in (Author, Title) and custom properties, edit them via WorkbookMetadata with DocumentProperties options, save to a new file, and verify that the changes are retained.
// Keywords: Aspose.Cells | C# | WorkbookMetadata | DocumentProperties | built‑in document properties | custom document properties | modify Excel metadata | verify Excel file properties | metadata API | save without full workbook
// Common Searches: Aspose.Cells change author property C# | Add custom document property with WorkbookMetadata | Read and write Excel metadata without opening workbook | Persist modified document properties in .xlsx using Aspose | WorkbookMetadata save updated properties
// Developer Intent: Edit built‑in and custom document properties of an Excel file through WorkbookMetadata and confirm the updates are saved.
// Use Cases: Set Author and Title, add a custom property, then modify them via WorkbookMetadata and save to a new workbook. | Load an existing .xlsx, change a built‑in property, add another custom property, and persist changes without re‑saving the whole workbook. | After updating metadata, reopen the file to read back all properties and ensure they match the expected values.
// AI Prompts: Generate C# code that opens an existing .xlsx, updates the Author to "Jane Doe" using WorkbookMetadata, adds a custom property "ReviewedOn" with today’s date, and saves to a new file. | Explain how WorkbookMetadata with MetadataOptions.DocumentProperties enables read‑write access to both built‑in and custom Excel properties without loading the full workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Metadata;
using Aspose.Cells.Properties;

namespace AsposeCellsMetadataDemo
{
    // Demonstrates how to create an Excel workbook, set built‑in (Author, Title) and custom properties, edit them via WorkbookMetadata with DocumentProperties options, save to a new file, and verify that the changes are retained.
    class Program
    {
        static void Main()
        {
            // Paths for the original and the updated workbook
            string originalPath = "Original.xlsx";
            string updatedPath = "Updated.xlsx";

            // ------------------------------------------------------------
            // 1. Create a new workbook and set built‑in & custom properties
            // ------------------------------------------------------------
            Workbook workbook = new Workbook();

            // Set some built‑in document properties
            workbook.BuiltInDocumentProperties.Author = "John Smith";
            workbook.BuiltInDocumentProperties.Title = "Metadata Demo";

            // Add a custom document property
            workbook.CustomDocumentProperties.Add("Project", "Aspose.Cells Metadata");

            // Save the workbook to disk (uses the standard Workbook.Save method)
            workbook.Save(originalPath);

            // ------------------------------------------------------------
            // 2. Load the workbook metadata (document properties) using WorkbookMetadata
            // ------------------------------------------------------------
            // Create MetadataOptions to work with document properties
            MetadataOptions options = new MetadataOptions(MetadataType.DocumentProperties);

            // Load metadata from the saved workbook (uses the WorkbookMetadata(string, MetadataOptions) constructor)
            WorkbookMetadata metadata = new WorkbookMetadata(originalPath, options);

            // Access built‑in properties (read‑write)
            BuiltInDocumentPropertyCollection builtInProps = metadata.BuiltInDocumentProperties;
            Console.WriteLine("Built‑in Author (before): " + builtInProps.Author);
            Console.WriteLine("Built‑in Title (before): " + builtInProps.Title);

            // Modify a built‑in property via metadata
            builtInProps.Author = "Aspose Developer";

            // Access custom properties
            CustomDocumentPropertyCollection customProps = metadata.CustomDocumentProperties;
            Console.WriteLine("Custom Property 'Project': " + customProps["Project"].Value);

            // Add another custom property
            customProps.Add("ReviewedBy", "Jane Doe");

            // ------------------------------------------------------------
            // 3. Save the modified metadata to a new file
            // ------------------------------------------------------------
            // Save using the WorkbookMetadata.Save(string) method
            metadata.Save(updatedPath);

            // ------------------------------------------------------------
            // 4. Verify that the properties are persisted in the saved file
            // ------------------------------------------------------------
            Workbook verifiedWorkbook = new Workbook(updatedPath);

            // Verify built‑in properties
            Console.WriteLine("Verified Built‑in Author: " + verifiedWorkbook.BuiltInDocumentProperties.Author);
            Console.WriteLine("Verified Built‑in Title: " + verifiedWorkbook.BuiltInDocumentProperties.Title);

            // Verify custom properties
            Console.WriteLine("Verified Custom Property 'Project': " + verifiedWorkbook.CustomDocumentProperties["Project"].Value);
            Console.WriteLine("Verified Custom Property 'ReviewedBy': " + verifiedWorkbook.CustomDocumentProperties["ReviewedBy"].Value);
        }
    }
}
