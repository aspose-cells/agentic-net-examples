// Title: Validate an XML map against an XSD schema with Aspose.Cells for .NET and display detailed validation errors
// AI Prompts: Write C# code that loads an Excel workbook using Aspose.Cells, builds an XmlSchemaSet from an XSD file, and runs schema validation on a target XML document via XmlReaderSettings, capturing all warnings and errors. | Show how to attach a ValidationEventHandler to XmlReaderSettings to capture schema validation messages and store them in a List<string> for later reporting. | Modify the validation flow to abort on the first error and throw an exception containing the error details instead of aggregating messages.
// Common Searches: c# aspnet cells validate xml against xsd and get error list | how to use XmlReaderSettings with Aspose.Cells to validate an XML map | console app to validate XML file with XSD schema and display warnings in .NET | read XSD into XmlSchemaSet and validate external XML using Aspose.Cells example
// Tags: aspocells xml validation using XmlReaderSettings | xml schema validation with XmlSchemaSet in C# | collect xml validation errors into List<string> | console output of xml validation results | load workbook before performing xml validation with Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Xml.Schema;

// The example loads an Excel workbook with Aspose.Cells, reads an XSD schema into an XmlSchemaSet, configures XmlReaderSettings for schema validation, validates a specified XML file, gathers any validation warnings or errors into a list, and writes the validation outcome to the console.
class Program
{
    static void Main()
    {
        try
        {
            // Paths to required files
            const string workbookPath = "input.xlsx";
            const string xmlFilePath = "data.xml";
            const string xsdFilePath = "schema.xsd";

            // Verify workbook file exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook file not found: {workbookPath}");
                return;
            }

            // Load the workbook (required for potential future use)
            Workbook workbook = new Workbook(workbookPath);

            // Verify XSD schema file exists
            if (!File.Exists(xsdFilePath))
            {
                Console.WriteLine($"XSD schema file not found: {xsdFilePath}");
                return;
            }

            string xsdContent;
            try
            {
                xsdContent = File.ReadAllText(xsdFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to read XSD file: {ex.Message}");
                return;
            }

            // Verify XML file exists
            if (!File.Exists(xmlFilePath))
            {
                Console.WriteLine($"XML file not found: {xmlFilePath}");
                return;
            }

            // Prepare an XmlSchemaSet and add the XSD schema
            XmlSchemaSet schemaSet = new XmlSchemaSet();
            using (StringReader sr = new StringReader(xsdContent))
            using (XmlReader xr = XmlReader.Create(sr))
            {
                schemaSet.Add(null, xr);
            }

            // Configure XmlReaderSettings for schema validation
            XmlReaderSettings settings = new XmlReaderSettings
            {
                // Resolve ambiguous reference by using fully qualified enum
                ValidationType = System.Xml.ValidationType.Schema,
                Schemas = schemaSet,
                ValidationFlags = XmlSchemaValidationFlags.ReportValidationWarnings
            };

            // Collect validation errors
            List<string> validationErrors = new List<string>();
            settings.ValidationEventHandler += (sender, e) =>
            {
                validationErrors.Add($"{e.Severity}: {e.Message}");
            };

            // Perform validation by reading the XML file
            try
            {
                using (XmlReader reader = XmlReader.Create(xmlFilePath, settings))
                {
                    while (reader.Read()) { }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during XML validation: {ex.Message}");
                return;
            }

            // Output validation results
            if (validationErrors.Count == 0)
            {
                Console.WriteLine("XML validation succeeded. No errors found.");
            }
            else
            {
                Console.WriteLine("XML validation failed with the following errors:");
                foreach (string error in validationErrors)
                {
                    Console.WriteLine(error);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
