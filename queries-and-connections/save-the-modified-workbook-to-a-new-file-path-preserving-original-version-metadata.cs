using System;
using Aspose.Cells;
using Aspose.Cells.Metadata;

namespace AsposeCellsMetadataPreserveDemo
{
    class Program
    {
        static void Main()
        {
            // Paths for the original and the new workbook
            string originalPath = "OriginalWorkbook.xlsx";
            string newPath = "ModifiedWorkbook.xlsx";

            // Load the original workbook
            Workbook workbook = new Workbook(originalPath);

            // ----- Example modification -----
            // Change the value of cell A1 in the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Modified content");
            // --------------------------------

            // Save the modified workbook content to the new file
            // (Uses the Workbook.Save(string) rule)
            workbook.Save(newPath, SaveFormat.Xlsx);

            // Preserve the original version metadata (document properties)
            // Create MetadataOptions for document properties
            MetadataOptions metaOptions = new MetadataOptions(MetadataType.DocumentProperties);

            // Load metadata from the original workbook
            // (Uses the WorkbookMetadata(string, MetadataOptions) constructor rule)
            WorkbookMetadata originalMetadata = new WorkbookMetadata(originalPath, metaOptions);

            // Save the metadata to the new workbook file
            // (Uses the WorkbookMetadata.Save(string) rule)
            originalMetadata.Save(newPath);

            Console.WriteLine("Workbook modified and saved with original metadata preserved.");
        }
    }
}