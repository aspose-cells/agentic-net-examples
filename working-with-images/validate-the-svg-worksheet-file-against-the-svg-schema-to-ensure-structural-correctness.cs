// Title: Validate Worksheet SVG Against XSD Schema Using Aspose.Cells for .NET
// Description: Shows how to generate a worksheet SVG with SheetRender, load an external SVG XSD into an XmlSchemaSet, and validate the SVG file using XmlReaderSettings and a ValidationEventHandler, outputting any schema errors.
// Keywords: Aspose.Cells | SVG validation | C# XSD | XmlSchemaSet | SheetRender | SVG export | XML schema validation .NET | validate SVG file | worksheet to SVG | SVG XSD schema
// Common Searches: Aspose.Cells validate SVG | C# validate SVG with XSD | how to check SVG schema in .NET | validate exported worksheet SVG | SVG schema validation example C#
// Developer Intent: Confirm that the SVG generated from a worksheet conforms to the SVG XSD specification.
// Use Cases: Run validation after exporting a worksheet to SVG to guarantee compliance with the SVG standard before publishing. | Add SVG schema checks to a CI/CD pipeline to detect rendering regressions early. | Capture detailed validation messages for debugging issues in worksheet‑to‑SVG conversion.
// AI Prompts: Write C# code that loads an XSD schema and validates an existing SVG file, returning all validation errors. | Explain which SvgImageOptions settings affect SVG compliance with the official SVG XSD. | Provide a step‑by‑step guide to integrate SVG XSD validation into an automated build process.

using System;
using System.IO;
using System.Xml;
using System.Xml.Schema;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Shows how to generate a worksheet SVG with SheetRender, load an external SVG XSD into an XmlSchemaSet, and validate the SVG file using XmlReaderSettings and a ValidationEventHandler, outputting any schema errors.
class SvgValidationDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Cells["A1"].PutValue("Item");
            worksheet.Cells["B1"].PutValue("Quantity");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("Orange");
            worksheet.Cells["B3"].PutValue(20);

            // Render the worksheet to an SVG file
            string svgFilePath = "worksheet.svg";
            SvgImageOptions svgOptions = new SvgImageOptions
            {
                FitToViewPort = true
            };
            SheetRender renderer = new SheetRender(worksheet, svgOptions);
            renderer.ToImage(0, svgFilePath);

            // Verify that the SVG file was created
            if (!File.Exists(svgFilePath))
            {
                Console.WriteLine($"Failed to create SVG file at '{svgFilePath}'.");
                return;
            }

            // Path to the SVG XSD schema file (must exist on disk)
            string schemaFilePath = "svg.xsd";

            // Ensure the schema file exists
            if (!File.Exists(schemaFilePath))
            {
                Console.WriteLine($"Schema file not found: '{schemaFilePath}'.");
                return;
            }

            // Prepare schema set
            XmlSchemaSet schemaSet = new XmlSchemaSet();
            try
            {
                schemaSet.Add(null, schemaFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading schema: {ex.Message}");
                return;
            }

            // Configure XML reader settings for validation
            XmlReaderSettings settings = new XmlReaderSettings
            {
                ValidationType = System.Xml.ValidationType.Schema,
                Schemas = schemaSet
            };

            bool isValid = true;
            settings.ValidationEventHandler += (sender, e) =>
            {
                Console.WriteLine($"Validation {e.Severity}: {e.Message}");
                isValid = false;
            };

            // Perform validation by reading the SVG file
            try
            {
                using (XmlReader reader = XmlReader.Create(svgFilePath, settings))
                {
                    while (reader.Read()) { }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during validation: {ex.Message}");
                return;
            }

            Console.WriteLine(isValid
                ? "SVG file is valid against the schema."
                : "SVG file is NOT valid.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
