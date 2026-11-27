// Title: C# example: Validate XML against an XSD schema, report errors, and import the XML as an Aspose.Cells XML map into an Excel workbook
// Common Searches: C# how to validate an XML file against an XSD schema before using Aspose.Cells | Aspose.Cells import XML data into Excel after schema validation in .NET | Report XML validation errors with XmlReader and XmlSchemaSet in C# | Add an XML map from XSD to a workbook using Aspose.Cells C# API | Save a validated workbook as XLSX after importing XML with Aspose.Cells
// Tags: xml validation with xsd using .net xmlreader | aspose.cells create xml map from xsd schema | import xml into worksheet aspnet cells | capture xml schema validation errors c# | export validated workbook to xlsx

using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Xml.Schema;
using Aspose.Cells;

namespace XmlMapValidationDemo
{
    // The sample program validates an XML file against a supplied XSD schema using .NET XmlReader and XmlSchemaSet, collects any validation errors, and aborts if the XML is invalid. When validation succeeds, it creates a new Aspose.Cells Workbook, adds the XSD as an XML map, imports the XML data into Sheet1, and saves the workbook as an XLSX file.
    class Program
    {
        // List to hold validation error messages
        private static readonly List<string> ValidationErrors = new List<string>();

        static void Main(string[] args)
        {
            // Expect three arguments: path to XSD, path to XML, optional output workbook path
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: XmlMapValidationDemo <XsdPath> <XmlPath> [OutputWorkbookPath]");
                return;
            }

            string xsdPath = args[0];
            string xmlPath = args[1];
            string outputWorkbookPath = args.Length >= 3 ? args[2] : "ValidatedWorkbook.xlsx";

            // Verify files exist
            if (!File.Exists(xsdPath))
            {
                Console.WriteLine($"XSD file not found: {xsdPath}");
                return;
            }

            if (!File.Exists(xmlPath))
            {
                Console.WriteLine($"XML file not found: {xmlPath}");
                return;
            }

            // Step 1: Validate XML against XSD using .NET XmlReader
            Console.WriteLine("Validating XML against XSD...");
            bool isValid = ValidateXmlAgainstXsd(xmlPath, xsdPath);

            if (!isValid)
            {
                Console.WriteLine("Validation failed. Errors:");
                foreach (var err in ValidationErrors)
                {
                    Console.WriteLine($"  - {err}");
                }
                // Stop further processing because the XML does not conform to the schema
                return;
            }

            Console.WriteLine("XML is valid according to the XSD schema.");

            // Step 2: Load workbook and add the XML map (the XSD) to the workbook
            Workbook workbook = new Workbook();

            try
            {
                // Add the XML map from the XSD file
                int mapIndex = workbook.Worksheets.XmlMaps.Add(xsdPath);
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];

                Console.WriteLine($"XML map added. Map name: {xmlMap.Name}, Root element: {xmlMap.RootElementName}");

                // Step 3: Import the XML data into the workbook using the map
                workbook.ImportXml(xmlPath, "Sheet1", 0, 0);
                Console.WriteLine("XML data imported into workbook successfully.");

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputWorkbookPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputWorkbookPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to: {outputWorkbookPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while working with Aspose.Cells: {ex.Message}");
            }
        }

        /// <param name="xmlFilePath">Path to the XML file.</param>
        /// <param name="xsdFilePath">Path to the XSD file.</param>
        /// <returns>True if XML is valid; otherwise false.</returns>
        private static bool ValidateXmlAgainstXsd(string xmlFilePath, string xsdFilePath)
        {
            ValidationErrors.Clear();

            XmlReaderSettings settings = new XmlReaderSettings
            {
                // Disambiguate the ValidationType enum from System.Xml
                ValidationType = System.Xml.ValidationType.Schema,
                DtdProcessing = DtdProcessing.Prohibit
            };

            // Load the XSD schema
            try
            {
                XmlSchemaSet schemas = new XmlSchemaSet();
                using (FileStream xsdStream = File.OpenRead(xsdFilePath))
                {
                    schemas.Add(null, XmlReader.Create(xsdStream));
                }
                settings.Schemas = schemas;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load XSD schema: {ex.Message}");
                return false;
            }

            // Attach validation event handler
            settings.ValidationEventHandler += ValidationCallback;

            // Parse the XML document
            try
            {
                using (FileStream xmlStream = File.OpenRead(xmlFilePath))
                using (XmlReader reader = XmlReader.Create(xmlStream, settings))
                {
                    while (reader.Read())
                    {
                        // Reading triggers validation
                    }
                }
            }
            catch (XmlException xmlEx)
            {
                ValidationErrors.Add($"XML parsing error: {xmlEx.Message}");
                return false;
            }
            catch (Exception ex)
            {
                ValidationErrors.Add($"Unexpected error during validation: {ex.Message}");
                return false;
            }

            // If any errors were collected, validation failed
            return ValidationErrors.Count == 0;
        }

        private static void ValidationCallback(object sender, ValidationEventArgs args)
        {
            string severity = args.Severity == XmlSeverityType.Error ? "Error" : "Warning";
            ValidationErrors.Add($"{severity}: {args.Message}");
        }
    }
}
