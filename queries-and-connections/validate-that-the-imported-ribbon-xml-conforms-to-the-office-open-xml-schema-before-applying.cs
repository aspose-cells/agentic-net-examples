// Title: Validate a Ribbon XML file against its XSD schema using C# XmlReaderSettings
// Common Searches: C# validate Ribbon XML file against its XSD schema before loading | How to set up XmlReaderSettings for Office Open XML ribbon validation in .NET | Log schema validation warnings and errors when checking custom UI XML with C#
// Tags: XmlReaderSettings ribbon custom UI schema validation | customUI.xsd validation using C# XmlReader | Office Open XML ribbon validation example | C# XML schema validation event handling | validate Ribbon custom UI against XSD

using System;
using System.IO;
using System.Xml;
using System.Xml.Schema;

// The RibbonXmlValidator class uses XmlReaderSettings with a loaded XSD schema to validate a Ribbon XML file, reporting errors via a ValidationEventHandler and returning true only when the XML conforms to the Office Open XML ribbon schema.
public class RibbonXmlValidator
{
    // Holds validation errors
    private static readonly ValidationEventHandler ValidationHandler = (sender, args) =>
    {
        Console.WriteLine($"Validation {args.Severity}: {args.Message}");
    };

    /// <param name="ribbonXmlPath">Full path to the Ribbon XML file to validate.</param>
    /// <param name="schemaPath">Full path to the Ribbon XSD schema file (customUI.xsd).</param>
    /// <returns>True if the XML conforms to the schema; otherwise false.</returns>
    public static bool Validate(string ribbonXmlPath, string schemaPath)
    {
        if (!File.Exists(ribbonXmlPath))
        {
            Console.WriteLine($"Ribbon XML file not found: {ribbonXmlPath}");
            return false;
        }

        if (!File.Exists(schemaPath))
        {
            Console.WriteLine($"Schema file not found: {schemaPath}");
            return false;
        }

        // Prepare the settings with the schema
        XmlReaderSettings settings = new XmlReaderSettings();
        settings.ValidationType = ValidationType.Schema;
        settings.ValidationFlags |= XmlSchemaValidationFlags.ProcessInlineSchema;
        settings.ValidationFlags |= XmlSchemaValidationFlags.ReportValidationWarnings;
        settings.ValidationEventHandler += ValidationHandler;

        // Load the schema
        XmlSchemaSet schemas = new XmlSchemaSet();
        schemas.Add(null, schemaPath);
        settings.Schemas = schemas;

        bool isValid = true;

        // Parse the XML with validation enabled
        using (XmlReader reader = XmlReader.Create(ribbonXmlPath, settings))
        {
            try
            {
                while (reader.Read()) { /* reading triggers validation */ }
            }
            catch (XmlException ex)
            {
                Console.WriteLine($"XML parsing error: {ex.Message}");
                isValid = false;
            }
            catch (XmlSchemaValidationException ex)
            {
                Console.WriteLine($"Schema validation error: {ex.Message}");
                isValid = false;
            }
        }

        // If any validation warnings/errors were raised, the handler will have printed them.
        // The handler does not change a flag, so we rely on exceptions for fatal errors.
        // For a stricter check, you could set a flag inside the handler.

        return isValid;
    }

    // Example usage
    public static void Main()
    {
        // Paths to the Ribbon XML and its XSD schema.
        // Adjust these paths to match your environment.
        string ribbonXml = @"C:\Path\To\customUI.xml";
        string ribbonXsd = @"C:\Path\To\customUI.xsd";

        bool result = Validate(ribbonXml, ribbonXsd);
        Console.WriteLine($"Ribbon XML validation result: {result}");
    }
}
