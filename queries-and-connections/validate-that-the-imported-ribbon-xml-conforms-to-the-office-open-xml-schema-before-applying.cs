using System;
using System.IO;
using System.Xml;
using System.Xml.Schema;
using Aspose.Cells;

namespace AsposeCellsRibbonXmlValidation
{
    class Program
    {
        // Simple XSD for the custom UI schema (subset for demonstration)
        private const string CustomUiSchema = @"
        <xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'
                   targetNamespace='http://schemas.microsoft.com/office/2006/01/customui'
                   xmlns='http://schemas.microsoft.com/office/2006/01/customui'
                   elementFormDefault='qualified'>

          <xs:element name='customUI'>
            <xs:complexType>
              <xs:sequence>
                <xs:element name='ribbon' minOccurs='0' maxOccurs='1'>
                  <xs:complexType>
                    <xs:sequence>
                      <xs:element name='tabs' minOccurs='0' maxOccurs='1'>
                        <xs:complexType>
                          <xs:sequence>
                            <xs:element name='tab' minOccurs='0' maxOccurs='unbounded'>
                              <xs:complexType>
                                <xs:sequence>
                                  <xs:element name='group' minOccurs='0' maxOccurs='unbounded'>
                                    <xs:complexType>
                                      <xs:sequence>
                                        <xs:element name='button' minOccurs='0' maxOccurs='unbounded' />
                                      </xs:sequence>
                                      <xs:attribute name='id' type='xs:string' use='required' />
                                      <xs:attribute name='label' type='xs:string' use='required' />
                                    </xs:complexType>
                                  </xs:element>
                                </xs:sequence>
                                <xs:attribute name='id' type='xs:string' use='required' />
                                <xs:attribute name='label' type='xs:string' use='required' />
                              </xs:complexType>
                            </xs:element>
                          </xs:sequence>
                        </xs:complexType>
                      </xs:element>
                    </xs:sequence>
                  </xs:complexType>
                </xs:element>
              </xs:sequence>
            </xs:complexType>
          </xs:element>

        </xs:schema>";

        static void Main()
        {
            try
            {
                // Sample Ribbon XML to be validated
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

                // Validate the XML against the schema
                bool isValid = ValidateXml(ribbonXml, CustomUiSchema, out string validationMessage);

                if (!isValid)
                {
                    Console.WriteLine("Ribbon XML validation failed:");
                    Console.WriteLine(validationMessage);
                    return;
                }

                // Create a new workbook (lifecycle rule)
                Workbook workbook = new Workbook();

                // Apply the validated Ribbon XML
                workbook.RibbonXml = ribbonXml;

                // Save the workbook (lifecycle rule)
                string outputPath = "ValidatedRibbonWorkbook.xlsm";
                workbook.Save(outputPath);

                Console.WriteLine($"Workbook saved successfully with validated Ribbon XML at '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected error occurred:");
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// Validates an XML string against a provided XSD schema string.
        /// </summary>
        /// <param name="xmlContent">The XML to validate.</param>
        /// <param name="xsdContent">The XSD schema.</param>
        /// <param name="errorMessage">Detailed validation errors, if any.</param>
        /// <returns>True if XML is valid; otherwise false.</returns>
        private static bool ValidateXml(string xmlContent, string xsdContent, out string errorMessage)
        {
            bool isValid = true;
            StringWriter errors = new StringWriter();

            // Prepare schema set
            XmlSchemaSet schemas = new XmlSchemaSet();
            using (StringReader xsdReader = new StringReader(xsdContent))
            {
                schemas.Add("http://schemas.microsoft.com/office/2006/01/customui",
                            XmlReader.Create(xsdReader));
            }

            // Configure XML reader settings for validation
            XmlReaderSettings settings = new XmlReaderSettings
            {
                ValidationType = System.Xml.ValidationType.Schema, // Resolve ambiguity
                Schemas = schemas,
                ValidationFlags =
                    XmlSchemaValidationFlags.ProcessIdentityConstraints |
                    XmlSchemaValidationFlags.ReportValidationWarnings
            };
            settings.ValidationEventHandler += (sender, args) =>
            {
                isValid = false;
                errors.WriteLine($"{args.Severity}: {args.Message}");
            };

            // Perform validation
            using (StringReader xmlReader = new StringReader(xmlContent))
            using (XmlReader reader = XmlReader.Create(xmlReader, settings))
            {
                try
                {
                    while (reader.Read()) { } // Parse entire document
                }
                catch (XmlException ex)
                {
                    isValid = false;
                    errors.WriteLine($"XML Exception: {ex.Message}");
                }
            }

            errorMessage = errors.ToString();
            return isValid;
        }
    }
}