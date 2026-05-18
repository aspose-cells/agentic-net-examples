using System;
using System.IO;
using System.Xml;
using System.Xml.Schema;
using Aspose.Cells;

namespace AsposeCellsRibbonXmlValidation
{
    class Program
    {
        // Sample Ribbon XML to be validated and applied
        private const string RibbonXml =
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

        // Minimal Custom UI schema (for demonstration purposes)
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
            // Validate the Ribbon XML against the Custom UI schema
            bool isValid = ValidateXml(RibbonXml, CustomUiSchema, out string validationMessage);

            if (!isValid)
            {
                Console.WriteLine("Ribbon XML validation failed:");
                Console.WriteLine(validationMessage);
                return;
            }

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Apply the validated Ribbon XML
            workbook.RibbonXml = RibbonXml;

            // Save the workbook (as macro-enabled to keep Ribbon UI)
            workbook.Save("ValidatedRibbonWorkbook.xlsm");

            Console.WriteLine("Workbook saved successfully with validated Ribbon XML.");
        }

        /// <summary>
        /// Validates an XML string against an XSD schema string.
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
            using (StringReader sr = new StringReader(xsdContent))
            {
                schemas.Add("http://schemas.microsoft.com/office/2006/01/customui", XmlReader.Create(sr));
            }

            // Configure XML reader settings for validation
            XmlReaderSettings settings = new XmlReaderSettings
            {
                ValidationType = System.Xml.ValidationType.Schema,
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
                    while (reader.Read()) { }
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