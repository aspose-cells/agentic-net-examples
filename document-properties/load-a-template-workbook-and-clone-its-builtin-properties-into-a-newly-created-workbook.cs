// Title: Clone Built‑In Document Properties from a Template Workbook with Aspose.Cells for .NET (C#)
// Description: This C# sample loads a template Excel file (or creates one), reads its built‑in metadata collection, copies each entry to a fresh Workbook, and saves the file, illustrating property cloning with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | clone built‑in document properties | copy Excel metadata | Workbook property transfer | Aspose.Cells .NET example
// Common Searches: copy Excel built‑in properties using Aspose.Cells | how to transfer workbook metadata in C# | Aspose.Cells clone document properties from template | C# code to duplicate Excel file metadata | Aspose.Cells copy built‑in document properties to new workbook
// Developer Intent: Duplicate all built‑in metadata from a source Excel workbook to a newly created workbook.
// Use Cases: Standardize report files by inheriting author, company, and category metadata from a master template. | Generate compliant spreadsheets with identical metadata for regulatory reporting. | Refresh legacy Excel files by recreating them while preserving original built‑in properties. | Automate batch processing where each new workbook must carry the same template metadata.
// AI Prompts: Generate C# code that uses Aspose.Cells to copy every built‑in document property from one workbook to another and saves the destination file. | Create a reusable function `CopyBuiltInProperties(string sourcePath, string destPath)` with exception handling and logging. | Explain how to extend the snippet to also copy custom document properties and preserve property types.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCloneBuiltInProperties
{
    // This C# sample loads a template Excel file (or creates one), reads its built‑in metadata collection, copies each entry to a fresh Workbook, and saves the file, illustrating property cloning with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                const string templatePath = "TemplateWorkbook.xlsx";
                const string destPath = "ClonedPropertiesWorkbook.xlsx";

                Workbook sourceWorkbook;

                // Load the template workbook if it exists; otherwise create a new one
                if (File.Exists(templatePath))
                {
                    sourceWorkbook = new Workbook(templatePath);
                }
                else
                {
                    sourceWorkbook = new Workbook();
                    // Example built‑in property for the generated source workbook
                    sourceWorkbook.BuiltInDocumentProperties["Author"].Value = "Default Author";
                }

                // Destination workbook (empty)
                Workbook destWorkbook = new Workbook();

                // Clone all built‑in document properties from source to destination
                foreach (var sourceProp in sourceWorkbook.BuiltInDocumentProperties)
                {
                    // The destination workbook already contains the same built‑in properties
                    destWorkbook.BuiltInDocumentProperties[sourceProp.Name].Value = sourceProp.Value;
                }

                // Save the destination workbook
                destWorkbook.Save(destPath);
                Console.WriteLine($"Workbook saved to '{destPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
