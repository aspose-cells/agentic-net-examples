using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Xml.Schema;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsSvgValidationDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Step 1: Create a simple workbook with sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Item");
                sheet.Cells["B1"].PutValue("Quantity");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("Orange");
                sheet.Cells["B3"].PutValue(15);

                // Step 2: Render the worksheet to an SVG file
                string svgPath = "worksheet.svg";
                SvgImageOptions svgOptions = new SvgImageOptions
                {
                    // ImageType is implicitly SVG for SvgImageOptions; no need to set it explicitly
                    FitToViewPort = true
                };
                SheetRender renderer = new SheetRender(sheet, svgOptions);
                renderer.ToImage(0, svgPath);

                // Verify that the SVG file was created
                if (!File.Exists(svgPath))
                {
                    Console.WriteLine($"Failed to generate SVG file: {svgPath}");
                    return;
                }

                // Step 3: Validate the generated SVG against an SVG XSD schema
                // Assume the SVG schema file (svg.xsd) is placed in the same directory as the executable
                string xsdPath = "svg.xsd";
                if (!File.Exists(xsdPath))
                {
                    Console.WriteLine($"Schema file not found: {xsdPath}");
                    return;
                }

                // Collect validation errors
                List<string> validationErrors = new List<string>();
                XmlReaderSettings settings = new XmlReaderSettings();
                settings.Schemas.Add(null, xsdPath);
                settings.ValidationType = System.Xml.ValidationType.Schema;
                settings.ValidationEventHandler += (sender, args) =>
                {
                    validationErrors.Add(args.Message);
                };

                // Perform validation
                using (XmlReader reader = XmlReader.Create(svgPath, settings))
                {
                    try
                    {
                        while (reader.Read()) { }
                    }
                    catch (XmlException ex)
                    {
                        validationErrors.Add($"XML parsing error: {ex.Message}");
                    }
                }

                // Step 4: Report validation result
                if (validationErrors.Count == 0)
                {
                    Console.WriteLine("SVG file is valid against the schema.");
                }
                else
                {
                    Console.WriteLine("SVG validation failed with the following errors:");
                    foreach (string error in validationErrors)
                    {
                        Console.WriteLine("- " + error);
                    }
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions to prevent the program from crashing
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}