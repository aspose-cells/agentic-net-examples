using System;
using Aspose.Cells;
using Aspose.Cells.Metadata;
using Aspose.Cells.Properties;

namespace CustomDocumentPropertiesDemo
{
    class Program
    {
        static void Main()
        {
            // Step 1: Create a simple workbook and save it.
            Workbook workbook = new Workbook();                     // create a new workbook
            workbook.Worksheets[0].Cells["A1"].PutValue("Demo");   // add sample data
            string originalPath = "Original.xlsx";
            workbook.Save(originalPath);                            // save the workbook

            // Step 2: Prepare metadata options to work with document properties.
            MetadataOptions options = new MetadataOptions(MetadataType.DocumentProperties);

            // Step 3: Load the workbook metadata from the saved file.
            WorkbookMetadata metadata = new WorkbookMetadata(originalPath, options);

            // Step 4: Add custom document properties of various types.
            metadata.CustomDocumentProperties.Add("ProjectName", "AsposeDemo");          // string
            metadata.CustomDocumentProperties.Add("Revision", 3);                       // int
            metadata.CustomDocumentProperties.Add("GeneratedOn", DateTime.UtcNow);     // DateTime
            metadata.CustomDocumentProperties.Add("IsApproved", true);                 // bool

            // Step 5: Save the metadata (which updates the workbook) to a new file.
            string outputPath = "WithCustomProps.xlsx";
            metadata.Save(outputPath);   // save the modified metadata

            // Optional verification: load the workbook and read back the custom properties.
            Workbook verifiedWorkbook = new Workbook(outputPath);
            Console.WriteLine("Custom Document Properties in the saved workbook:");
            foreach (DocumentProperty prop in verifiedWorkbook.CustomDocumentProperties)
            {
                Console.WriteLine($"{prop.Name}: {prop.Value} ({prop.Type})");
            }
        }
    }
}