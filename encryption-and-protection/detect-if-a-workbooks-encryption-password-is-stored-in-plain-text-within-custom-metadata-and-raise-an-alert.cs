using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Metadata;

class DetectPlainTextPasswordInMetadata
{
    static void Main()
    {
        // Path to the workbook file
        string filePath = "SensitiveWorkbook.xlsx";

        try
        {
            // -----------------------------------------------------------------
            // Step 1: Create a workbook, set an encryption password, add a custom
            //         document property that (insecurely) stores the same password,
            //         and save the file.
            // -----------------------------------------------------------------
            using (Workbook wb = new Workbook())
            {
                // Set workbook encryption password
                wb.Settings.Password = "Secret123";

                // Add some sample data
                wb.Worksheets[0].Cells["A1"].PutValue("Sample Data");

                // Save the workbook (encrypted)
                wb.Save(filePath);
            }

            // Ensure the file was created before working with metadata
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"The workbook file '{filePath}' was not found.");

            // Open metadata for the saved workbook to add a custom property
            MetadataOptions metaOptions = new MetadataOptions(MetadataType.DocumentProperties);
            WorkbookMetadata metadata = new WorkbookMetadata(filePath, metaOptions);

            // Insecurely store the encryption password in a custom property
            metadata.CustomDocumentProperties.Add("EncryptionPassword", "Secret123");

            // Save metadata changes back to the workbook
            metadata.Save(filePath);

            // -----------------------------------------------------------------
            // Step 2: Load the workbook (using the known password) to retrieve the
            //         actual encryption password from the workbook settings.
            // -----------------------------------------------------------------
            LoadOptions loadOpts = new LoadOptions { Password = "Secret123" };
            using (Workbook loadedWb = new Workbook(filePath, loadOpts))
            {
                string workbookPassword = loadedWb.Settings.Password; // should be "Secret123"

                // -----------------------------------------------------------------
                // Step 3: Load the custom metadata and check if any property value
                //         matches the workbook's encryption password.
                // -----------------------------------------------------------------
                WorkbookMetadata loadedMeta = new WorkbookMetadata(filePath, metaOptions);

                foreach (var prop in loadedMeta.CustomDocumentProperties)
                {
                    // Only consider string values
                    if (prop.Value is string value && value == workbookPassword)
                    {
                        Console.WriteLine(
                            $"ALERT: Encryption password is stored in plain text in custom metadata property '{prop.Name}'.");
                    }
                }
            }
        }
        catch (FileNotFoundException fnfEx)
        {
            Console.WriteLine($"File error: {fnfEx.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}