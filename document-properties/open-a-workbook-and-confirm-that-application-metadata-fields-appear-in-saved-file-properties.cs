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
            // Define file paths
            string originalPath = "OriginalWorkbook.xlsx";
            string updatedPath = "UpdatedWorkbook.xlsx";

            // -------------------------------------------------
            // 1. Create a new workbook and save it (lifecycle: create & save)
            // -------------------------------------------------
            Workbook workbook = new Workbook();                     // create workbook
            workbook.Worksheets[0].Cells["A1"].PutValue("Demo");   // add sample data
            workbook.Save(originalPath);                            // save workbook

            // -------------------------------------------------
            // 2. Load workbook metadata for document properties (lifecycle: load)
            // -------------------------------------------------
            MetadataOptions options = new MetadataOptions(MetadataType.DocumentProperties);
            WorkbookMetadata metadata = new WorkbookMetadata(originalPath, options); // load metadata

            // -------------------------------------------------
            // 3. Modify built‑in and custom document properties
            // -------------------------------------------------
            // Built‑in properties (read/write)
            metadata.BuiltInDocumentProperties.Author = "Aspose Developer";
            metadata.BuiltInDocumentProperties.Title = "Metadata Demo";

            // Custom properties (add new)
            metadata.CustomDocumentProperties.Add("Project", "WorkbookMetadataDemo");
            metadata.CustomDocumentProperties.Add("Version", 1);

            // -------------------------------------------------
            // 4. Save the modified metadata back to a new file (lifecycle: save)
            // -------------------------------------------------
            metadata.Save(updatedPath);

            // -------------------------------------------------
            // 5. Load the updated workbook and verify properties
            // -------------------------------------------------
            Workbook updatedWorkbook = new Workbook(updatedPath);

            // Verify built‑in properties
            string author = updatedWorkbook.BuiltInDocumentProperties["Author"].Value.ToString();
            string title = updatedWorkbook.BuiltInDocumentProperties["Title"].Value.ToString();

            // Verify custom properties
            string project = updatedWorkbook.CustomDocumentProperties["Project"].Value.ToString();
            string version = updatedWorkbook.CustomDocumentProperties["Version"].Value.ToString();

            // Output verification results
            Console.WriteLine("Verification of saved metadata:");
            Console.WriteLine($"Author (built‑in): {author}");
            Console.WriteLine($"Title (built‑in): {title}");
            Console.WriteLine($"Project (custom): {project}");
            Console.WriteLine($"Version (custom): {version}");
        }
    }
}