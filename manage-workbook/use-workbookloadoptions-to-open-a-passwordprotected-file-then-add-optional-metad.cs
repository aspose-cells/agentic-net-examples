using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Metadata;

namespace AsposeCellsPasswordProtectedExample
{
    class Program
    {
        static void Main()
        {
            // Path to the existing password‑protected workbook
            string protectedFilePath = "protected.xlsx";

            // Password used to open the workbook
            string password = "myPassword";

            // Load the workbook using LoadOptions with the password
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = password;
            Workbook workbook = new Workbook(protectedFilePath, loadOptions);

            // (Optional) Modify the workbook – e.g., write a value to a cell
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Data after opening with password");

            // Save the workbook to a memory stream (required for metadata manipulation)
            using (MemoryStream stream = new MemoryStream())
            {
                workbook.Save(stream, SaveFormat.Xlsx);
                stream.Position = 0; // Reset stream position for reading

                // Create MetadataOptions for document properties and set the same password
                MetadataOptions metaOptions = new MetadataOptions(MetadataType.DocumentProperties);
                metaOptions.Password = password;

                // Load metadata from the stream using the options
                WorkbookMetadata metadata = new WorkbookMetadata(stream, metaOptions);

                // Add a custom document property (optional metadata)
                metadata.CustomDocumentProperties.Add("ReviewedBy", "John Doe");
                metadata.CustomDocumentProperties.Add("ReviewDate", DateTime.Now);

                // Save the workbook together with the updated metadata to a new file
                string outputFilePath = "protected_with_metadata.xlsx";
                metadata.Save(outputFilePath);
            }

            // Clean up
            workbook.Dispose();

            Console.WriteLine("Workbook opened with password, metadata added, and saved successfully.");
        }
    }
}