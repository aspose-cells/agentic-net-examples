using System;
using Aspose.Cells;
using Aspose.Cells.Metadata;

namespace AsposeCellsMetadataPreserveDemo
{
    public class Program
    {
        public static void Main()
        {
            // Paths for the original workbook and the new modified workbook
            string sourcePath = "original.xlsx";
            string destinationPath = "modified.xlsx";

            // Load the original workbook
            Workbook workbook = new Workbook(sourcePath);

            // Example modification: change the value of cell A1 in the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Modified Content");

            // Save the modified workbook to the new file path
            workbook.Save(destinationPath, SaveFormat.Xlsx);

            // Preserve the original version metadata (document properties)
            // Create MetadataOptions for document properties
            MetadataOptions metaOptions = new MetadataOptions(MetadataType.DocumentProperties);

            // Load metadata from the original workbook
            WorkbookMetadata metadata = new WorkbookMetadata(sourcePath, metaOptions);

            // Save the metadata to the newly created workbook file
            metadata.Save(destinationPath);

            Console.WriteLine("Workbook saved with original metadata preserved.");
        }
    }
}