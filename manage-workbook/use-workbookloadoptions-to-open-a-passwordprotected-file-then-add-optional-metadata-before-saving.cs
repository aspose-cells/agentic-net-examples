// Title: Open a password‑protected Excel workbook with LoadOptions and add a custom document property using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an encrypted .xlsx file with Aspose.Cells LoadOptions, adds a custom document property via WorkbookMetadata, and saves the file. | Show how to create a protected workbook if it doesn't exist, open it with the correct password, modify its metadata, and handle common exceptions.
// Common Searches: Aspose.Cells load encrypted workbook C# | add custom document property to protected Excel file Aspose.Cells | WorkbookMetadata password protected .xlsx | save metadata changes after opening secured workbook .NET
// Tags: Aspose.Cells | LoadOptions | password protected workbook | WorkbookMetadata | custom document property | C# | exception handling

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Metadata;

// The sample creates a password‑protected Excel file if needed, opens it with LoadOptions using the correct password, loads its document‑property metadata, adds a custom property, and saves both the metadata and the workbook while handling file‑not‑found, Aspose.Cells, and generic exceptions.
class LoadProtectedWorkbookWithMetadata
{
    static void Main()
    {
        // Path to the password‑protected workbook
        string filePath = "ProtectedWorkbook.xlsx";
        string password = "mySecretPwd";

        try
        {
            // If the workbook does not exist, create a simple one and protect it with a password
            if (!File.Exists(filePath))
            {
                var wb = new Workbook();
                wb.Worksheets[0].Cells["A1"].PutValue("Sample Data");
                // Set the password required to open the workbook
                wb.Settings.Password = password;
                wb.Save(filePath);
                Console.WriteLine("Sample protected workbook created.");
            }

            // Load the protected workbook using LoadOptions with the correct password
            var loadOptions = new LoadOptions
            {
                Password = password
            };
            var workbook = new Workbook(filePath, loadOptions);

            // Verify that the workbook was loaded correctly
            Console.WriteLine("Loaded cell value: " + workbook.Worksheets[0].Cells["A1"].Value);

            // Load workbook metadata with the same password
            var metaOptions = new MetadataOptions(MetadataType.DocumentProperties)
            {
                Password = password
            };
            var metadata = new WorkbookMetadata(filePath, metaOptions);

            // Add a custom document property
            metadata.CustomDocumentProperties.Add("ReviewedBy", "John Doe");

            // Save the updated metadata back to the file
            metadata.Save(filePath);

            // Save the workbook (if any worksheet changes were made)
            workbook.Save(filePath);

            Console.WriteLine("Workbook opened, metadata added, and saved successfully.");
        }
        catch (FileNotFoundException fnfEx)
        {
            Console.WriteLine("File not found: " + fnfEx.Message);
        }
        catch (CellsException cellsEx)
        {
            Console.WriteLine("Aspose.Cells error: " + cellsEx.Message);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Unexpected error: " + ex.Message);
        }
    }
}
