using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Metadata;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        try
        {
            const string sourceFile = "source.xlsx";
            const string destFile = "dest.xlsx";

            // Verify source file exists to avoid FileNotFoundException
            if (!File.Exists(sourceFile))
            {
                Console.WriteLine($"Source file \"{sourceFile}\" not found.");
                return;
            }

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourceFile);

            // Define the source range to copy (e.g., A1:B5)
            AsposeRange sourceRange = sourceWorkbook.Worksheets[0].Cells.CreateRange("A1:B5");

            // Create a new (empty) destination workbook
            Workbook destinationWorkbook = new Workbook();

            // Define the destination range with the same size
            AsposeRange destinationRange = destinationWorkbook.Worksheets[0].Cells.CreateRange("A1:B5");

            // Copy the source range into the destination range
            sourceRange.Copy(destinationRange);

            // Save the destination workbook to a file
            destinationWorkbook.Save(destFile);

            // Set the author property using document metadata
            MetadataOptions options = new MetadataOptions(MetadataType.DocumentProperties);
            WorkbookMetadata metadata = new WorkbookMetadata(destFile, options);
            metadata.BuiltInDocumentProperties.Author = "John Doe";
            metadata.Save(destFile);

            Console.WriteLine($"Workbook copied and saved to \"{destFile}\" with author metadata set.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}