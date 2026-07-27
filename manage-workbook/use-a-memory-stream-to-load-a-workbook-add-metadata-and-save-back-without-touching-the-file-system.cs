using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Metadata;

class Program
{
    static void Main()
    {
        // 1. Create a workbook entirely in memory and add some data.
        Workbook wb = new Workbook();
        wb.Worksheets[0].Cells["A1"].PutValue("Hello");
        wb.Worksheets[0].Cells["B1"].PutValue(DateTime.Now);

        // 2. Save the workbook to a memory stream (XLSX format).
        using (MemoryStream wbStream = new MemoryStream())
        {
            wb.Save(wbStream, SaveFormat.Xlsx);
            wbStream.Position = 0; // Reset for reading.

            // 3. Prepare metadata options to work with document properties.
            MetadataOptions metaOptions = new MetadataOptions(MetadataType.DocumentProperties);

            // 4. Load the workbook metadata from the memory stream.
            WorkbookMetadata metadata = new WorkbookMetadata(wbStream, metaOptions);

            // 5. Add custom document properties.
            metadata.CustomDocumentProperties.Add("ProcessedBy", "AsposeDemo");
            metadata.CustomDocumentProperties.Add("ProcessedOn", DateTime.UtcNow);

            // 6. Save the modified metadata back to a new memory stream.
            using (MemoryStream metaOutStream = new MemoryStream())
            {
                metadata.Save(metaOutStream);
                metaOutStream.Position = 0; // Reset for reading.

                // 7. Load the workbook again from the metadata‑updated stream.
                Workbook finalWb = new Workbook(metaOutStream);

                // 8. Verify that the custom properties are present.
                Console.WriteLine("Custom Property 'ProcessedBy': " +
                    finalWb.CustomDocumentProperties["ProcessedBy"].Value);
                Console.WriteLine("Custom Property 'ProcessedOn': " +
                    finalWb.CustomDocumentProperties["ProcessedOn"].Value);

                // 9. Save the final workbook to another memory stream (no file system access).
                using (MemoryStream finalStream = new MemoryStream())
                {
                    finalWb.Save(finalStream, SaveFormat.Xlsx);
                    Console.WriteLine("Final workbook saved to memory stream, length: " + finalStream.Length);
                }
            }
        }
    }
}