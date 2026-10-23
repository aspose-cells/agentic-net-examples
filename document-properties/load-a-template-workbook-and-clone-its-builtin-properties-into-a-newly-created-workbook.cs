// Title: How to clone built‑in document properties from a template Excel workbook to a new workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells copy built‑in document properties from one workbook to another | how to clone Excel file metadata using Aspose.Cells .NET | transfer template workbook properties to a new workbook with Aspose.Cells | Aspose.Cells example for duplicating built‑in properties when creating a workbook | load a template Excel file and copy its document properties in C#
// Tags: copy built‑in properties Aspose.Cells | clone workbook metadata .NET | Aspose.Cells BuiltInDocumentProperties iteration | duplicate Excel template properties C# | transfer workbook metadata using Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The code loads a template workbook (or creates an empty one if missing), creates a new workbook, iterates through the template's BuiltInDocumentProperties collection, copies each property's value to the matching property in the new workbook, and saves the result.
class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "Template.xlsx";
            const string outputPath = "ClonedPropertiesWorkbook.xlsx";

            // Load the template workbook if it exists; otherwise use an empty workbook
            Workbook templateWorkbook;
            if (File.Exists(templatePath))
            {
                templateWorkbook = new Workbook(templatePath);
            }
            else
            {
                Console.WriteLine($"Template file '{templatePath}' not found. Using an empty workbook as source.");
                templateWorkbook = new Workbook();
            }

            // Create a new workbook where the properties will be cloned
            Workbook newWorkbook = new Workbook();

            // Clone each built‑in property from the template to the new workbook
            foreach (var prop in templateWorkbook.BuiltInDocumentProperties)
            {
                // Ensure the target property exists before assigning the value
                if (newWorkbook.BuiltInDocumentProperties.Contains(prop.Name))
                {
                    newWorkbook.BuiltInDocumentProperties[prop.Name].Value = prop.Value;
                }
            }

            // Save the newly created workbook with the cloned properties
            newWorkbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
