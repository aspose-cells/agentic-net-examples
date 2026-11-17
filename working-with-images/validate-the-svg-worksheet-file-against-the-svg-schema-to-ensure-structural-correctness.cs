// Title: Validate an SVG file against its XSD schema using C# and XmlReader
// Common Searches: C# code to validate an SVG file with an XSD schema | how to use XmlReaderSettings for SVG schema validation in .NET | checking SVG structural compliance with XSD in a console app | validate SVG against svg.xsd using XmlSchemaSet in C# | detect invalid SVG elements during XML parsing in C#
// Tags: C# XmlReader schema validation for SVG | XmlSchemaSet integration with SVG XSD | SVG XSD validation using .NET XmlReader | console app SVG structural validation | handling validation events for SVG in C#

using System;
using System.IO;
using System.Xml;
using System.Xml.Schema;

namespace SvgValidationApp
{
    // The example loads an SVG file and its XSD schema, ensures both exist, adds the schema to an XmlSchemaSet, configures XmlReaderSettings for schema validation, attaches a ValidationEventHandler to capture errors, reads the SVG with XmlReader, and reports whether the SVG conforms to the schema.
    class SvgValidator
    {
        static void Main()
        {
            try
            {
                // Paths to the SVG file and its XSD schema
                string svgFilePath = "input.svg";
                string svgSchemaPath = "svg.xsd";

                // Verify that required files exist
                if (!File.Exists(svgFilePath))
                {
                    Console.WriteLine($"SVG file not found: {svgFilePath}");
                    return;
                }

                if (!File.Exists(svgSchemaPath))
                {
                    Console.WriteLine($"SVG schema file not found: {svgSchemaPath}");
                    return;
                }

                // Load the SVG schema
                XmlSchemaSet schemaSet = new XmlSchemaSet();
                schemaSet.Add(null, svgSchemaPath);

                // Track validation result
                bool isValid = true;

                // Handle validation events
                ValidationEventHandler validationHandler = (sender, e) =>
                {
                    Console.WriteLine($"{e.Severity}: {e.Message}");
                    isValid = false;
                };

                // Configure XML reader for schema validation
                XmlReaderSettings settings = new XmlReaderSettings
                {
                    ValidationType = ValidationType.Schema,
                    Schemas = schemaSet
                };
                settings.ValidationEventHandler += validationHandler;

                // Read and validate the SVG file
                using (XmlReader reader = XmlReader.Create(svgFilePath, settings))
                {
                    while (reader.Read()) { }
                }

                // Output validation result
                Console.WriteLine(isValid ? "SVG is valid." : "SVG is invalid.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
