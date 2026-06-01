using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Metadata;

class Program
{
    static void Main()
    {
        // Path to the password‑protected workbook
        string filePath = "ProtectedWorkbook.xlsx";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Load the workbook with the required password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = "secure123"
            };
            Workbook workbook = new Workbook(filePath, loadOptions);

            // (Optional) modify workbook content
            workbook.Worksheets[0].Cells["A1"].PutValue("Demo content");

            // Prepare metadata options for document properties with the same password
            MetadataOptions metaOptions = new MetadataOptions(MetadataType.DocumentProperties)
            {
                Password = "secure123"
            };

            // Load workbook metadata
            WorkbookMetadata metadata = new WorkbookMetadata(filePath, metaOptions);

            // Add a custom document property
            metadata.CustomDocumentProperties.Add("Project", "AsposeDemo");

            // Save metadata changes back to the file
            metadata.Save(filePath);

            // Save any workbook content changes (overwrites the same file)
            workbook.Save(filePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}