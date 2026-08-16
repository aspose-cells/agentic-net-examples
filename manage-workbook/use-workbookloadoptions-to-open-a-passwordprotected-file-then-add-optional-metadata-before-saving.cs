// Title: Open a password‑protected Excel workbook with LoadOptions, add custom document properties via MetadataOptions, and save – Aspose.Cells for .NET
// Description: This example shows how to load an encrypted .xlsx file using LoadOptions (supplying the opening password), optionally unprotect it, add or update custom document properties through WorkbookMetadata and MetadataOptions (using the same password), and then save both the metadata and workbook changes while keeping the file protected.
// Keywords: Aspose.Cells | LoadOptions password | protected workbook | custom document properties | MetadataOptions | .NET Excel encryption | add metadata to encrypted file | WorkbookMetadata
// Common Searches: Aspose.Cells open password protected Excel file | LoadOptions opening password .NET | Add custom document properties to protected workbook | MetadataOptions password protected Excel | Update metadata in encrypted .xlsx using Aspose.Cells
// Developer Intent: Open a password‑protected workbook, modify its custom properties, and save the file with protection intact.
// Use Cases: Read an existing encrypted workbook, change or add custom document properties, and re‑save without removing the password. | Create a new workbook, set both opening and modification passwords, then embed custom metadata before the first save. | Unprotect a write‑protected sheet, write data to cells, update metadata, and save while re‑applying protection.
// AI Prompts: Write C# code that uses Aspose.Cells LoadOptions to open an encrypted Excel file, adds a custom property "ReviewedBy", and saves the workbook preserving its password. | Explain step‑by‑step how to use MetadataOptions with a password to add or update custom document properties in a protected workbook using Aspose.Cells for .NET. | Provide a concise example that creates a new workbook, applies opening and modification passwords, adds custom metadata, and saves the file.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Metadata;

namespace AsposeCellsPasswordProtectedExample
{
    // This example shows how to load an encrypted .xlsx file using LoadOptions (supplying the opening password), optionally unprotect it, add or update custom document properties through WorkbookMetadata and MetadataOptions (using the same password), and then save both the metadata and workbook changes while keeping the file protected.
    class Program
    {
        static void Main()
        {
            // Path to the workbook that will be password‑protected
            const string filePath = "ProtectedWorkbook.xlsx";
            const string password = "securePassword123";

            try
            {
                // If the file does not exist, create a new workbook and protect it with a password
                if (!File.Exists(filePath))
                {
                    var newWb = new Workbook();

                    // Set opening password (required to open)
                    newWb.Settings.Password = password;

                    // Set modification password (required to modify)
                    newWb.Protect(ProtectionType.All, password);

                    newWb.Save(filePath);
                }

                // Load the workbook using LoadOptions with the opening password
                var loadOptions = new LoadOptions { Password = password };
                var workbook = new Workbook(filePath, loadOptions);

                // If the workbook is write‑protected, unprotect it before making changes
                try
                {
                    workbook.Unprotect(password);
                }
                catch
                {
                    // Ignore if the workbook is not write‑protected
                }

                // Prepare metadata options (specify the password if the document is protected)
                var metaOptions = new MetadataOptions(MetadataType.DocumentProperties) { Password = password };
                var metadata = new WorkbookMetadata(filePath, metaOptions);

                // Add custom document properties
                metadata.CustomDocumentProperties.Add("ReviewedBy", "John Doe");
                metadata.CustomDocumentProperties.Add("ReviewDate", DateTime.Now);

                // Save the updated metadata back to the file
                metadata.Save(filePath);

                // Demonstrate a workbook change
                workbook.Worksheets[0].Cells["A1"].PutValue("Metadata updated");
                workbook.Save(filePath);

                Console.WriteLine("Workbook opened with password, metadata added, and file saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
