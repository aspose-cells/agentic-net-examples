using System;
using System.IO;
using System.Text;
using System.Xml;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    public class ExportVbaCertificateToXml
    {
        public static void Run()
        {
            // Load a workbook that contains a signed VBA project
            string workbookPath = "SignedWorkbook.xlsm";
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Verify that the VBA project is signed and certificate data is available
            if (vbaProject.IsSigned && vbaProject.CertRawData != null && vbaProject.CertRawData.Length > 0)
            {
                // Get the raw certificate bytes
                byte[] certBytes = vbaProject.CertRawData;

                // Convert the certificate bytes to a Base64 string for XML representation
                string certBase64 = Convert.ToBase64String(certBytes);

                // Create an XML document that holds the certificate data
                XmlDocument xmlDoc = new XmlDocument();

                // Create XML declaration
                XmlDeclaration xmlDecl = xmlDoc.CreateXmlDeclaration("1.0", "UTF-8", null);
                xmlDoc.AppendChild(xmlDecl);

                // Create root element
                XmlElement root = xmlDoc.CreateElement("VbaCertificate");
                xmlDoc.AppendChild(root);

                // Add a child element with the Base64-encoded certificate
                XmlElement certElement = xmlDoc.CreateElement("CertificateData");
                certElement.InnerText = certBase64;
                root.AppendChild(certElement);

                // Option 1: Save the XML to a file
                string xmlFilePath = "VbaCertificate.xml";
                xmlDoc.Save(xmlFilePath);
                Console.WriteLine($"Certificate exported to XML file: {xmlFilePath}");

                // Option 2: Save the XML to a memory stream
                using (MemoryStream xmlStream = new MemoryStream())
                {
                    // Write the XML document into the stream using an XmlWriter
                    XmlWriterSettings settings = new XmlWriterSettings
                    {
                        Encoding = Encoding.UTF8,
                        Indent = true,
                        CloseOutput = false
                    };

                    using (XmlWriter writer = XmlWriter.Create(xmlStream, settings))
                    {
                        xmlDoc.WriteTo(writer);
                        writer.Flush();
                    }

                    // Reset stream position for reading or further processing
                    xmlStream.Position = 0;

                    // Example: read back the XML as a string (optional)
                    using (StreamReader reader = new StreamReader(xmlStream, Encoding.UTF8, true, 1024, true))
                    {
                        string xmlContent = reader.ReadToEnd();
                        Console.WriteLine("Certificate exported to XML stream. Content preview:");
                        Console.WriteLine(xmlContent.Substring(0, Math.Min(200, xmlContent.Length)) + "...");
                    }

                    // The xmlStream can now be used wherever a Stream is required
                }
            }
            else
            {
                Console.WriteLine("The VBA project is not signed or certificate data is unavailable.");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            ExportVbaCertificateToXml.Run();
        }
    }
}