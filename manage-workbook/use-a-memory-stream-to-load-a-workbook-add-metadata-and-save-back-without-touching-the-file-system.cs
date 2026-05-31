using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Metadata;

namespace AsposeCellsMemoryMetadataDemo
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample");
            sheet.Cells["B1"].PutValue(123);

            // 2. Save the workbook to a memory stream (XLSX format)
            using (MemoryStream workbookStream = new MemoryStream())
            {
                workbook.Save(workbookStream, SaveFormat.Xlsx);
                workbookStream.Position = 0; // Reset for reading

                // 3. Prepare metadata options for document properties
                MetadataOptions metaOptions = new MetadataOptions(MetadataType.DocumentProperties);

                // 4. Load metadata from the workbook stream
                WorkbookMetadata metadata = new WorkbookMetadata(workbookStream, metaOptions);

                // 5. Add a custom document property
                metadata.CustomDocumentProperties.Add("MyCustomProperty", "CustomValue");

                // 6. Save the modified metadata to a new memory stream
                using (MemoryStream updatedStream = new MemoryStream())
                {
                    metadata.Save(updatedStream);
                    updatedStream.Position = 0; // Reset for reading

                    // 7. Load a workbook from the updated stream to verify the metadata
                    Workbook resultWorkbook = new Workbook(updatedStream);

                    // 8. Output the custom property value to the console
                    var prop = resultWorkbook.CustomDocumentProperties["MyCustomProperty"];
                    Console.WriteLine($"Custom Property 'MyCustomProperty' = {prop?.Value}");
                }
            }

            // Clean up
            workbook.Dispose();
        }
    }
}