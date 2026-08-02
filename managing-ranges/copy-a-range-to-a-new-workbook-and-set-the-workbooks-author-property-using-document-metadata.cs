// Title: Copy an Excel range to a new workbook and set the Author property with Aspose.Cells for .NET
// Description: Load a source workbook, copy a defined cell range (e.g., A1:B5) into a fresh workbook, save it, then use WorkbookMetadata to assign the built‑in Author property before the final save. Demonstrates range copying and document metadata handling in C#.
// Keywords: Aspose.Cells copy range | C# copy Excel cells | Aspose.Cells WorkbookMetadata | set Author property Excel | document properties Aspose.Cells | copy range to new workbook .NET | Excel metadata update
// Common Searches: Aspose.Cells copy range to another workbook C# | How to set Author metadata in an Excel file using Aspose.Cells | Copy cells between Excel files and update document properties | WorkbookMetadata example for setting built‑in properties | Save Excel file after copying range with Aspose
// Developer Intent: Copy a specific cell range into a newly created workbook and assign an Author value in the workbook’s built‑in document properties.
// Use Cases: Create a report workbook by extracting a data block from a template and embedding the author’s name for audit purposes. | Automate generation of personalized workbooks where each file contains copied content and the creator’s identifier stored in metadata. | Migrate selected data from legacy spreadsheets to clean files while ensuring compliance by programmatically setting document properties.
// AI Prompts: Generate C# code with Aspose.Cells that copies a named range from a source workbook to a destination workbook and sets multiple built‑in properties (Author, Title, Subject) before saving. | Explain how WorkbookMetadata works in Aspose.Cells for updating document properties after a workbook has been saved. | Write a reusable method that takes source path, range address, destination path, and author name, then copies the range and updates the Author property in the new file.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Metadata;
using AsposeRange = Aspose.Cells.Range;

// Load a source workbook, copy a defined cell range (e.g., A1:B5) into a fresh workbook, save it, then use WorkbookMetadata to assign the built‑in Author property before the final save. Demonstrates range copying and document metadata handling in C#.
class CopyRangeAndSetAuthor
{
    static void Main()
    {
        try
        {
            string sourcePath = "source.xlsx";
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file '{sourcePath}' not found.");
                return;
            }

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Define the source range to copy (e.g., A1:B5)
            AsposeRange sourceRange = sourceWorkbook.Worksheets[0].Cells.CreateRange("A1:B5");

            // Create a new (empty) destination workbook
            Workbook destWorkbook = new Workbook();

            // Define the destination range with the same size
            AsposeRange destRange = destWorkbook.Worksheets[0].Cells.CreateRange("A1:B5");

            // Copy the source range into the destination range
            sourceRange.Copy(destRange);

            // Save the destination workbook (initial save before metadata update)
            string destPath = "dest.xlsx";
            destWorkbook.Save(destPath);

            // Create metadata options for document properties
            MetadataOptions metaOptions = new MetadataOptions(MetadataType.DocumentProperties);

            // Load the metadata for the saved workbook
            WorkbookMetadata metadata = new WorkbookMetadata(destPath, metaOptions);

            // Set the built‑in Author property
            metadata.BuiltInDocumentProperties.Author = "John Doe";

            // Save the metadata back to the workbook file
            metadata.Save(destPath);

            Console.WriteLine("Range copied and author metadata set successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
