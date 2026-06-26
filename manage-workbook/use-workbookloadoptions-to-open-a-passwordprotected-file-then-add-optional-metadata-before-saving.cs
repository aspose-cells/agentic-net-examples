using System;
using Aspose.Cells;
using Aspose.Cells.Metadata;

namespace AsposeCellsPasswordProtectedMetadataDemo
{
    public class Program
    {
        public static void Main()
        {
            // Path to the existing password‑protected workbook
            string sourcePath = "protected.xlsx";

            // Load the workbook using LoadOptions with the required password
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = "myPassword";
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // (Optional) Modify workbook content if needed
            // workbook.Worksheets[0].Cells["A1"].PutValue("Updated value");

            // Prepare metadata options for document properties and set the same password
            MetadataOptions metaOptions = new MetadataOptions(MetadataType.DocumentProperties);
            metaOptions.Password = "myPassword";

            // Load metadata from the workbook using the metadata options
            WorkbookMetadata metadata = new WorkbookMetadata(sourcePath, metaOptions);

            // Add a custom document property
            metadata.CustomDocumentProperties.Add("ReviewedBy", "John Doe");

            // Save the workbook together with the updated metadata to a new file
            string outputPath = "protected_with_metadata.xlsx";
            metadata.Save(outputPath);

            // Verify that the file is still password‑protected by loading it again
            Workbook verifiedWorkbook = new Workbook(outputPath, loadOptions);
            Console.WriteLine("Verification successful. Cell A1 value: " + verifiedWorkbook.Worksheets[0].Cells["A1"].Value);
        }
    }
}