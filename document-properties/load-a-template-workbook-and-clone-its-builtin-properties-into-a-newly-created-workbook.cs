// Title: How to clone built‑in document properties from a template Excel workbook to a new workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells copy built‑in document properties from one Excel file to another | clone template workbook metadata to a new workbook with Aspose.Cells .NET | transfer Excel built‑in properties using Aspose.Cells API in C# | copy workbook built‑in document properties programmatically with Aspose.Cells | Aspose.Cells example for duplicating built‑in document properties in a new workbook
// Tags: copy built‑in document properties Aspose.Cells C# | clone Excel workbook metadata Aspose.Cells .NET | transfer built‑in properties between workbooks C# | Aspose.Cells built‑in document properties assignment | duplicate template workbook properties Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads a template Excel file, creates a new empty workbook, iterates over the template's BuiltInDocumentProperties, assigns each existing property's value to the matching property in the new workbook, and saves the new workbook with the cloned properties.
class Program
{
    static void Main()
    {
        const string templatePath = "Template.xlsx";
        const string outputPath = "ClonedWorkbook.xlsx";

        // Verify that the template file exists to avoid FileNotFoundException
        if (!File.Exists(templatePath))
        {
            Console.WriteLine($"Error: Template file not found at '{templatePath}'.");
            return;
        }

        try
        {
            // Load the template workbook
            Workbook templateWb = new Workbook(templatePath);

            // Create a new empty workbook
            Workbook newWb = new Workbook();

            // Clone built‑in document properties from the template to the new workbook
            foreach (dynamic srcProp in templateWb.BuiltInDocumentProperties)
            {
                // The target workbook already contains all built‑in properties,
                // so we can directly assign the value if the property exists.
                if (newWb.BuiltInDocumentProperties.Contains(srcProp.Name))
                {
                    newWb.BuiltInDocumentProperties[srcProp.Name].Value = srcProp.Value;
                }
            }

            // Save the new workbook with the cloned properties
            newWb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
