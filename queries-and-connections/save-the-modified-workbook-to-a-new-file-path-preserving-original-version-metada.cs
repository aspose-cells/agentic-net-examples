using System;
using Aspose.Cells;
using Aspose.Cells.Metadata;

namespace AsposeCellsMetadataPreserveDemo
{
    class Program
    {
        static void Main()
        {
            // Paths for the original workbook and the new modified workbook
            string originalPath = "original.xlsx";
            string modifiedPath = "modified.xlsx";

            // Load the original workbook
            Workbook workbook = new Workbook(originalPath);

            // Example modification: change the value of cell A1 in the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Modified Value");

            // Save the modified workbook to the new file path
            workbook.Save(modifiedPath, SaveFormat.Xlsx);

            // Preserve the original version metadata by copying it to the new file
            // Create MetadataOptions for document properties (you can include other types if needed)
            MetadataOptions metaOptions = new MetadataOptions(MetadataType.DocumentProperties);

            // Load the metadata from the original workbook
            WorkbookMetadata originalMetadata = new WorkbookMetadata(originalPath, metaOptions);

            // Save the metadata to the modified workbook file
            originalMetadata.Save(modifiedPath);

            Console.WriteLine("Workbook modified and saved with original metadata preserved.");
        }
    }
}