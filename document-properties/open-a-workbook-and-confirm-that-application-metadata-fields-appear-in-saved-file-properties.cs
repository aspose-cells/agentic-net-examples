// Title: Set and Verify Built‑in & Custom Document Properties with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, assign built‑in properties (Author, Title) and a custom property, save the file, load it with WorkbookMetadata using the DocumentProperties option, modify the Author, and confirm that both built‑in and custom properties persist after saving.
// Keywords: Aspose.Cells document properties | C# set built‑in Excel properties | add custom property Aspose.Cells | WorkbookMetadata modify author | verify Excel metadata .NET | save workbook with metadata | Aspose.Cells DocumentProperties example
// Common Searches: how to add custom document properties in Excel using Aspose.Cells C# | modify built‑in author property with WorkbookMetadata | verify Excel file metadata after saving Aspose.Cells | Aspose.Cells set built‑in properties programmatically | read custom properties from .xlsx with Aspose.Cells
// Developer Intent: Add and edit built‑in and custom document properties via WorkbookMetadata and ensure they are retained in the saved Excel file.
// Use Cases: Embed author and project information into a template workbook before distribution. | Update the Author field of generated reports without loading the full workbook, using WorkbookMetadata for performance. | Automated validation that custom properties survive modifications and re‑saving of Excel files.
// AI Prompts: Generate C# code that reads all custom document properties from an existing Excel file using Aspose.Cells and prints each name and value. | Provide a step‑by‑step tutorial to change the Title built‑in property of a workbook with WorkbookMetadata and save the result to a new file.

using System;
using Aspose.Cells;
using Aspose.Cells.Metadata;
using Aspose.Cells.Properties;

namespace AsposeCellsMetadataDemo
{
    // Demonstrates how to create a workbook, assign built‑in properties (Author, Title) and a custom property, save the file, load it with WorkbookMetadata using the DocumentProperties option, modify the Author, and confirm that both built‑in and custom properties persist after saving.
    class Program
    {
        static void Main()
        {
            // Paths for the intermediate and final files
            string originalPath = "Original.xlsx";
            string finalPath = "WithMetadata.xlsx";

            // -------------------------------------------------
            // 1. Create a new workbook and set built‑in properties
            // -------------------------------------------------
            Workbook workbook = new Workbook();

            // Set some built‑in document properties
            workbook.BuiltInDocumentProperties.Author = "John Doe";
            workbook.BuiltInDocumentProperties.Title = "Metadata Demo";

            // Add a custom document property
            workbook.CustomDocumentProperties.Add("Project", "Aspose.Cells");

            // Save the workbook to disk (uses the provided Save(string) rule)
            workbook.Save(originalPath);

            // -------------------------------------------------
            // 2. Load the workbook metadata with DocumentProperties option
            // -------------------------------------------------
            MetadataOptions options = new MetadataOptions(MetadataType.DocumentProperties);
            WorkbookMetadata metadata = new WorkbookMetadata(originalPath, options); // uses the provided constructor rule

            // (Optional) modify metadata here if needed, e.g. change built‑in author
            metadata.BuiltInDocumentProperties.Author = "Jane Smith";

            // Save the modified metadata to a new file (uses the provided Save(string) rule)
            metadata.Save(finalPath);

            // -------------------------------------------------
            // 3. Verify that the properties are present in the saved file
            // -------------------------------------------------
            Workbook loadedWorkbook = new Workbook(finalPath);

            // Retrieve and display built‑in properties
            Console.WriteLine("Built‑in Properties:");
            Console.WriteLine("Author: " + loadedWorkbook.BuiltInDocumentProperties.Author);
            Console.WriteLine("Title : " + loadedWorkbook.BuiltInDocumentProperties.Title);

            // Retrieve and display custom properties
            Console.WriteLine("\nCustom Properties:");
            foreach (DocumentProperty prop in loadedWorkbook.CustomDocumentProperties)
            {
                Console.WriteLine($"{prop.Name}: {prop.Value}");
            }
        }
    }
}
