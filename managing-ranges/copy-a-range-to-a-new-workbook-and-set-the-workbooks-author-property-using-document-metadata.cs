// Title: Copy a Range to a New Workbook and Set the Author Property with Aspose.Cells for .NET
// Description: Loads Source.xlsx, copies the A1:C5 range into a fresh workbook, saves it as Destination.xlsx, then updates the built‑in Author document property using Aspose.Cells metadata APIs. Demonstrates range transfer and metadata manipulation in C#.
// Keywords: Aspose.Cells copy range | C# copy cells between workbooks | Aspose.Cells set author property | Excel document metadata .NET | Workbook built‑in properties | range transfer Aspose.Cells | Aspose.Cells MetadataOptions | programmatic Excel author tag
// Common Searches: how to copy a range from one Excel file to another using Aspose.Cells | set author property in Excel workbook with Aspose.Cells .NET | Aspose.Cells copy cells and edit document properties | update Excel metadata without reopening the workbook | C# Aspose.Cells copy range and set built‑in properties
// Developer Intent: Transfer a specific cell block to a new workbook and assign an Author value via document metadata.
// Use Cases: Create a report workbook by extracting a data block from a template and embedding the report author’s name for compliance. | Automate generation of a summary file that contains selected cells from a source sheet while programmatically adding creator information. | Migrate a defined range to a separate file and set built‑in properties such as Author to track ownership and version history.
// AI Prompts: Generate C# code with Aspose.Cells that copies range A1:C5 from Source.xlsx to Destination.xlsx and sets the Author property to a given name. | Show how to copy a dynamic range between two workbooks and then update multiple built‑in document properties (Author, Title, Subject) using Aspose.Cells metadata. | Explain a method to copy a range and modify Excel metadata without reopening the source workbook in Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Metadata;
using AsposeRange = Aspose.Cells.Range;

// Loads Source.xlsx, copies the A1:C5 range into a fresh workbook, saves it as Destination.xlsx, then updates the built‑in Author document property using Aspose.Cells metadata APIs. Demonstrates range transfer and metadata manipulation in C#.
class Program
{
    static void Main()
    {
        // Paths for source and destination workbooks
        string sourcePath = "Source.xlsx";
        string destPath = "Destination.xlsx";

        try
        {
            // Ensure source file exists; if not, create a simple workbook
            if (!File.Exists(sourcePath))
            {
                var tempWb = new Workbook();
                var tempSheet = tempWb.Worksheets[0];
                tempSheet.Cells["A1"].PutValue("Sample");
                tempWb.Save(sourcePath);
            }

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Define the source range to copy (e.g., A1:C5)
            AsposeRange sourceRange = sourceWorkbook.Worksheets[0].Cells.CreateRange("A1:C5");

            // Create a new (empty) destination workbook
            Workbook destinationWorkbook = new Workbook();

            // Get the first worksheet of the destination workbook
            Worksheet destSheet = destinationWorkbook.Worksheets[0];

            // Create a destination range of the same size starting at A1
            AsposeRange destRange = destSheet.Cells.CreateRange("A1:C5");

            // Copy the source range into the destination range
            sourceRange.Copy(destRange);

            // Save the destination workbook
            destinationWorkbook.Save(destPath);

            // ---- Set the author property using document metadata ----
            // Create metadata options for document properties
            MetadataOptions options = new MetadataOptions(MetadataType.DocumentProperties);

            // Load metadata for the saved workbook
            WorkbookMetadata metadata = new WorkbookMetadata(destPath, options);

            // Set the built‑in Author property
            metadata.BuiltInDocumentProperties.Author = "John Doe";

            // Save the updated metadata back to the file
            metadata.Save(destPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
