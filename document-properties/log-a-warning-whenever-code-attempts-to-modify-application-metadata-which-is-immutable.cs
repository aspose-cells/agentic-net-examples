using System;
using Aspose.Cells;
using Aspose.Cells.Metadata;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // ------------------------------------------------------------
        // 1. Create a new workbook and save it to a physical file.
        // ------------------------------------------------------------
        Workbook workbook = new Workbook();
        string filePath = "Sample.xlsx";
        workbook.Save(filePath); // lifecycle rule: create -> save

        // ------------------------------------------------------------
        // 2. Prepare metadata options for loading document properties.
        // ------------------------------------------------------------
        MetadataOptions options = new MetadataOptions(MetadataType.DocumentProperties);

        // ------------------------------------------------------------
        // 3. Load the workbook metadata using the provided constructor.
        // ------------------------------------------------------------
        WorkbookMetadata metadata = new WorkbookMetadata(filePath, options); // lifecycle rule: load

        // ------------------------------------------------------------
        // 4. Attempt to modify application metadata (NameOfApplication).
        //    In this scenario the application metadata is considered immutable,
        //    so we log a warning before performing the change.
        // ------------------------------------------------------------
        Console.WriteLine("Warning: Attempting to modify immutable application metadata 'NameOfApplication'.");

        // Perform the modification (for demonstration purposes only).
        metadata.BuiltInDocumentProperties.NameOfApplication = "MyApp";

        // ------------------------------------------------------------
        // 5. Save the modified metadata back to the same file.
        // ------------------------------------------------------------
        metadata.Save(filePath); // lifecycle rule: save

        Console.WriteLine("Metadata saved successfully.");
    }
}