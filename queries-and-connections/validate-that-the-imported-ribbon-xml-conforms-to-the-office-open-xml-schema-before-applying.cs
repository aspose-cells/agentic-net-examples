using System;
using System.IO;
using System.Xml;
using System.Xml.Schema;
using Aspose.Cells;

class RibbonXmlValidationDemo
{
    static void Main()
    {
        try
        {
            // Sample Ribbon XML that will be applied to the workbook
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

            // Path to the Office Open XML Custom UI schema (XSD). 
            string schemaPath = "customUI.xsd";

            // Ensure the schema file exists before validation
            if (!File.Exists(schemaPath))
            {
                Console.WriteLine($"Schema file not found: {schemaPath}");
                return;
            }

            // Validate the Ribbon XML against the schema
            if (!ValidateXmlAgainstSchema(ribbonXml, schemaPath))
            {
                Console.WriteLine("Ribbon XML validation failed. The workbook will not be created.");
                return;
            }

            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Apply the validated Ribbon XML (feature rule: use RibbonXml property)
            workbook.RibbonXml = ribbonXml;

            // Save the workbook as a macro-enabled file to retain the Ribbon UI (lifecycle rule: save)
            string outputPath = "ValidatedRibbonWorkbook.xlsm";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully with validated Ribbon XML: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }

    // Helper method that validates an XML string against a given XSD file
    static bool ValidateXmlAgainstSchema(string xmlContent, string xsdFilePath)
    {
        bool isValid = true;

        // Prepare the schema set
        XmlSchemaSet schemaSet = new XmlSchemaSet();
        schemaSet.Add("http://schemas.microsoft.com/office/2006/01/customui", xsdFilePath);

        // Configure XML reader settings for schema validation
        XmlReaderSettings settings = new XmlReaderSettings
        {
            ValidationType = System.Xml.ValidationType.Schema,
            Schemas = schemaSet
        };
        settings.ValidationEventHandler += (sender, args) =>
        {
            Console.WriteLine($"Schema validation error: {args.Message}");
            isValid = false;
        };

        // Parse and validate the XML
        using (StringReader stringReader = new StringReader(xmlContent))
        using (XmlReader xmlReader = XmlReader.Create(stringReader, settings))
        {
            try
            {
                while (xmlReader.Read()) { }
            }
            catch (XmlException ex)
            {
                Console.WriteLine($"XML parsing error: {ex.Message}");
                isValid = false;
            }
        }

        return isValid;
    }
}