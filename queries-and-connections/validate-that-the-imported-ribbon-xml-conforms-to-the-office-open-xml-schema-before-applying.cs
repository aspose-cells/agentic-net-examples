// Title: Validate Office Ribbon XML against customUI.xsd and assign it to an Aspose.Cells workbook (C#)
// AI Prompts: Write C# code that loads a customUI.xsd file, validates a Ribbon XML string using XmlReaderSettings, and returns any validation errors. | Create a method that reads Ribbon XML from an external file, validates it against the Office Open XML schema, and sets Workbook.RibbonXml only when the XML passes validation. | Show how to capture and log the first XSD validation error while applying the verified Ribbon XML to a macro‑enabled .xlsm workbook with Aspose.Cells.
// Common Searches: c# validate customUI.xsd against ribbon xml prior to Aspose.Cells usage | console application schema check for Office ribbon UI and assign to Workbook.RibbonXml | .NET XmlReaderSettings example for Office custom UI XSD validation | export .xlsm file with verified ribbon XML using Aspose.Cells | read ribbon definition from external file and validate with customUI.xsd in C#
// Tags: xml schema validation with XmlReaderSettings | customUI.xsd Office ribbon validation | Aspose.Cells RibbonXml assignment | macro-enabled .xlsm workbook saving Aspose.Cells | C# load and validate ribbon XML

using System;
using System.IO;
using System.Xml;
using System.Xml.Schema;
using Aspose.Cells;

// Demonstrates loading the customUI.xsd schema, validating Ribbon XML with XmlReaderSettings, assigning the validated XML to Workbook.RibbonXml, and saving the workbook as a macro‑enabled .xlsm file using Aspose.Cells.
class ValidateRibbonXmlDemo
{
    static void Main()
    {
        try
        {
            // Sample Ribbon XML to be validated and applied
            string ribbonXml =
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
                "  <ribbon>" +
                "    <tabs>" +
                "      <tab id=\"customTab\" label=\"My Tab\">" +
                "        <group id=\"customGroup\" label=\"My Group\">" +
                "          <button id=\"customButton\" label=\"My Button\" size=\"large\" />" +
                "        </group>" +
                "      </tab>" +
                "    </tabs>" +
                "  </ribbon>" +
                "</customUI>";

            // Path to the Office Open XML Ribbon schema (customUI.xsd)
            string schemaPath = "customUI.xsd";

            // Ensure the schema file exists
            if (!File.Exists(schemaPath))
            {
                Console.WriteLine($"Schema file not found: {schemaPath}");
                return;
            }

            // Validate the Ribbon XML against the schema
            if (!ValidateXml(ribbonXml, schemaPath, out string validationMessage))
            {
                Console.WriteLine("Ribbon XML validation failed: " + validationMessage);
                return;
            }

            // Create a new workbook (standard creation rule)
            Workbook workbook = new Workbook();

            // Apply the validated Ribbon XML to the workbook
            workbook.RibbonXml = ribbonXml;

            // Save the workbook (macro-enabled format to retain Ribbon UI)
            string outputPath = "ValidatedRibbonWorkbook.xlsm";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook saved successfully with validated Ribbon XML: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An unexpected error occurred: " + ex.Message);
        }
    }

    // Helper method to validate XML string against an XSD schema
    static bool ValidateXml(string xmlContent, string schemaFilePath, out string message)
    {
        message = string.Empty;
        try
        {
            // Load the schema into a set
            XmlSchemaSet schemaSet = new XmlSchemaSet();
            schemaSet.Add("http://schemas.microsoft.com/office/2006/01/customui", schemaFilePath);

            // Configure XML reader settings for schema validation
            XmlReaderSettings settings = new XmlReaderSettings
            {
                ValidationType = System.Xml.ValidationType.Schema,
                Schemas = schemaSet
            };

            string validationError = null;
            settings.ValidationEventHandler += (sender, args) =>
            {
                // Capture the first validation error message
                if (validationError == null)
                {
                    validationError = args.Message;
                }
            };

            // Parse and validate the XML content
            using (StringReader stringReader = new StringReader(xmlContent))
            using (XmlReader reader = XmlReader.Create(stringReader, settings))
            {
                while (reader.Read()) { }
            }

            // Set output message based on validation result
            message = validationError ?? string.Empty;
            return string.IsNullOrEmpty(validationError);
        }
        catch (Exception ex)
        {
            message = ex.Message;
            return false;
        }
    }
}
