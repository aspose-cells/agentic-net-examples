// Title: Clone Workbook Metadata from a Template to a New File using Aspose.Cells (C#)
// Description: Shows how to load a template Excel file (or create a blank one), instantiate a fresh workbook, copy each built‑in property from the source to the target, and save the result with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | Excel metadata | built-in document properties | clone workbook properties | template Excel file | copy Excel metadata | document property API
// Common Searches: Aspose.Cells copy Excel built‑in properties C# | transfer workbook metadata with Aspose.Cells | clone template workbook properties .NET | programmatically set Excel document properties using C# | copy Excel file metadata with Aspose library
// Developer Intent: Copy the source workbook's metadata into a newly generated workbook.
// Use Cases: Generate reports that inherit author, title, and company information from a standard template. | Create batches of workbooks from a master file while preserving creation date, manager, and other metadata. | Automate archival copies of Excel files ensuring the original property values are retained.
// AI Prompts: Write C# code with Aspose.Cells that loads a template workbook, copies all built‑in document properties to a new workbook, and saves it. | Provide an Aspose.Cells .NET example that checks for a template file, falls back to an empty workbook if missing, and synchronizes their metadata fields. | Generate a reusable method that accepts two Workbook objects and mirrors the source's built‑in properties onto the destination.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

// Shows how to load a template Excel file (or create a blank one), instantiate a fresh workbook, copy each built‑in property from the source to the target, and save the result with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the template workbook (source)
            string templatePath = "TemplateWorkbook.xlsx";

            // Load the template workbook if it exists; otherwise create a new one
            Workbook sourceWorkbook;
            if (File.Exists(templatePath))
            {
                sourceWorkbook = new Workbook(templatePath);
            }
            else
            {
                Console.WriteLine($"Template file not found: {templatePath}. Creating a new source workbook.");
                sourceWorkbook = new Workbook(); // empty workbook as fallback
            }

            // Create a new empty workbook (destination)
            Workbook destWorkbook = new Workbook();

            // Clone built‑in document properties from source to destination
            foreach (DocumentProperty sourceProp in sourceWorkbook.BuiltInDocumentProperties)
            {
                // Ensure the destination has the same property name
                if (destWorkbook.BuiltInDocumentProperties.Contains(sourceProp.Name))
                {
                    destWorkbook.BuiltInDocumentProperties[sourceProp.Name].Value = sourceProp.Value;
                }
            }

            // Save the new workbook with cloned properties
            string outputPath = "ClonedPropertiesWorkbook.xlsx";
            destWorkbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
