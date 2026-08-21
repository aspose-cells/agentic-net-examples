// Title: Validate Ribbon XML Against the Office Open XML Schema Using Aspose.Cells (C#)
// Description: A C# console example that checks for the presence of a customUI.xml file, validates its content against the official Office Open XML custom UI schema with XmlSchemaSet, and, after successful validation, assigns the XML to Workbook.RibbonXml and saves the workbook as an .xlsm file.
// Keywords: Aspose.Cells ribbon validation | Office Open XML custom UI schema | C# XML schema validation | Workbook.RibbonXml | Excel custom ribbon .xlsm | XmlSchemaSet .NET | CI/CD ribbon XML check | Excel add‑in development
// Common Searches: how to validate ribbon xml with aspnet | c# validate customui.xml against office schema | aspocells set ribbon xml after validation | xmlschema validation for excel ribbon | embed validated ribbon xml in workbook using aspocells
// Developer Intent: Ensure a custom ribbon XML file conforms to the Office Open XML schema before embedding it in an Excel workbook with Aspose.Cells.
// Use Cases: Prevent runtime errors by verifying ribbon XML compliance before saving .xlsm workbooks. | Automate batch validation of multiple ribbon XML files in a CI pipeline. | Integrate strict XML validation into an Excel add‑in that generates workbooks with custom ribbon interfaces.
// AI Prompts: Create a reusable C# method that downloads and caches the Office custom UI schema, then validates any Ribbon XML string. | Write error‑logging code that captures validation details to a file instead of throwing exceptions, while still blocking invalid XML. | Refactor the validation routine to use async/await and support schema versions for Office 2016, 2019, and Microsoft 365.

using System;
using System.IO;
using System.Xml;
using System.Xml.Schema;
using Aspose.Cells;

namespace AsposeCellsRibbonValidation
{
    // A C# console example that checks for the presence of a customUI.xml file, validates its content against the official Office Open XML custom UI schema with XmlSchemaSet, and, after successful validation, assigns the XML to Workbook.RibbonXml and saves the workbook as an .xlsm file.
    class Program
    {
        // Path to the Ribbon XML file to be imported
        private const string RibbonXmlPath = "customUI.xml";

        // URL of the official Office Open XML custom UI schema
        private const string RibbonSchemaUrl = "http://schemas.microsoft.com/office/2006/01/customui";

        static void Main()
        {
            try
            {
                // Verify that the Ribbon XML file exists
                if (!File.Exists(RibbonXmlPath))
                {
                    Console.WriteLine($"Error: Ribbon XML file \"{RibbonXmlPath}\" not found.");
                    return;
                }

                // Load the Ribbon XML content
                string ribbonXml = File.ReadAllText(RibbonXmlPath);

                // Validate the XML against the Office Open XML schema
                ValidateRibbonXml(ribbonXml);

                // Create a new workbook (lifecycle rule)
                Workbook workbook = new Workbook();

                // Apply the validated Ribbon XML to the workbook
                workbook.RibbonXml = ribbonXml;

                // Save the workbook (lifecycle rule)
                string outputPath = "WorkbookWithRibbon.xlsm";
                workbook.Save(outputPath);

                Console.WriteLine($"Workbook saved successfully with validated Ribbon XML at \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        /// <param name="xmlContent">The Ribbon XML to validate.</param>
        private static void ValidateRibbonXml(string xmlContent)
        {
            // Prepare a schema set and add the custom UI schema from the official URL
            XmlSchemaSet schemaSet = new XmlSchemaSet();
            using (XmlReader schemaReader = XmlReader.Create(RibbonSchemaUrl))
            {
                schemaSet.Add(null, schemaReader);
            }

            // Configure XML reader settings for schema validation
            XmlReaderSettings settings = new XmlReaderSettings
            {
                ValidationType = System.Xml.ValidationType.Schema,
                Schemas = schemaSet,
                // Stop on the first validation error
                ValidationFlags = XmlSchemaValidationFlags.ReportValidationWarnings
            };
            settings.ValidationEventHandler += ValidationEventHandler;

            // Perform validation using an XmlReader over the XML string
            using (StringReader stringReader = new StringReader(xmlContent))
            using (XmlReader xmlReader = XmlReader.Create(stringReader, settings))
            {
                while (xmlReader.Read())
                {
                    // Reading triggers validation
                }
            }
        }

        // Handles validation events; throws on errors or warnings
        private static void ValidationEventHandler(object? sender, ValidationEventArgs e)
        {
            // Treat warnings as errors for strict validation
            throw new InvalidOperationException($"Ribbon XML validation error: {e.Message}");
        }
    }
}
