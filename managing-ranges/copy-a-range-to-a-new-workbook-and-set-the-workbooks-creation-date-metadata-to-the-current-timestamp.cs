using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Metadata;

class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "Source.xlsx";
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Define the source range to copy (A1:B5) on the first worksheet
            Aspose.Cells.Range sourceRange = sourceWorkbook.Worksheets[0].Cells.CreateRange("A1:B5");

            // Create a new (empty) workbook for the destination
            Workbook destinationWorkbook = new Workbook();

            // Define the destination range (C1:D5) on the first worksheet of the new workbook
            Aspose.Cells.Range destinationRange = destinationWorkbook.Worksheets[0].Cells.CreateRange("C1:D5");

            // Copy the source range to the destination range
            destinationRange.Copy(sourceRange);

            // Save the destination workbook
            const string destFilePath = "Destination.xlsx";
            destinationWorkbook.Save(destFilePath);

            // Update the workbook's creation date metadata to the current timestamp
            MetadataOptions metaOptions = new MetadataOptions(MetadataType.DocumentProperties);
            WorkbookMetadata metadata = new WorkbookMetadata(destFilePath, metaOptions);
            metadata.BuiltInDocumentProperties.CreatedTime = DateTime.Now;
            metadata.Save(destFilePath);

            Console.WriteLine($"Workbook successfully saved to {destFilePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}