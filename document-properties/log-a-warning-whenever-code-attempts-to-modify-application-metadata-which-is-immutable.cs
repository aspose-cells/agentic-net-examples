using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Metadata;

namespace AsposeCellsDemo
{
    public class MetadataImmutableDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // -------------------------------------------------
            // Create a simple workbook and persist it to disk
            // -------------------------------------------------
            Workbook workbook = new Workbook();
            workbook.Worksheets[0].Cells["A1"].Value = "Sample Data";

            string filePath = "ImmutableDemo.xlsx";

            try
            {
                workbook.Save(filePath);
                Console.WriteLine($"Workbook saved to '{filePath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
                return;
            }

            // -------------------------------------------------
            // Load workbook metadata for document properties
            // -------------------------------------------------
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            WorkbookMetadata metadata = null;
            try
            {
                MetadataOptions options = new MetadataOptions(MetadataType.DocumentProperties);
                metadata = new WorkbookMetadata(filePath, options);
                Console.WriteLine("Metadata loaded successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load metadata: {ex.Message}");
                return;
            }

            // -------------------------------------------------
            // Attempt to modify immutable application metadata
            // -------------------------------------------------
            Console.WriteLine("Warning: Attempting to modify immutable application metadata 'NameOfApplication'.");

            try
            {
                // This property may be immutable; Aspose.Cells will throw if modification is not allowed.
                metadata.BuiltInDocumentProperties.NameOfApplication = "MyCustomApp";
                Console.WriteLine("NameOfApplication property modified.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unable to modify NameOfApplication: {ex.Message}");
            }

            // -------------------------------------------------
            // Save the modified metadata back to the file
            // -------------------------------------------------
            try
            {
                metadata.Save(filePath);
                Console.WriteLine("Metadata saved back to the workbook.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save metadata: {ex.Message}");
            }
        }
    }
}