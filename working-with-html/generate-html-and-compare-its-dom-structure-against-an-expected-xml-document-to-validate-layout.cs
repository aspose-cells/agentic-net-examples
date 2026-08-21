// Title: Aspose.Cells C# – Validate generated HTML against an expected XML DOM
// Description: Creates an Excel workbook, saves it as HTML5 using Aspose.Cells HtmlSaveOptions, loads the output with XmlDocument, defines the target XML structure, and recursively compares both DOM trees to confirm that the HTML table layout, cell IDs and values match the specification. Ideal for automated regression testing of Excel‑to‑HTML conversions.
// Keywords: Aspose.Cells HTML5 export | C# HtmlSaveOptions | Excel to HTML validation | XmlDocument DOM comparison | recursive XML element match | HTML table layout verification | unit test for Aspose.Cells output | regression testing Excel HTML | C# compare generated HTML with XML | Aspose.Cells workbook to HTML
// Common Searches: compare Aspose.Cells generated HTML with XML in C# | validate HTML table produced by Aspose.Cells | C# DOM comparison for Excel to HTML conversion | how to test Aspose.Cells HTML output | recursive element comparison XmlDocument C#
// Developer Intent: Verify that the HTML produced by Aspose.Cells exactly mirrors a predefined XML representation of the desired page structure.
// Use Cases: Automated verification that each workbook cell appears as a <td> with the correct id and value in the exported HTML. | Regression test that flags any change in HtmlSaveOptions or worksheet data that alters the HTML hierarchy. | Pre‑deployment check ensuring the saved HTML complies with HTML5 standards and the expected DOM layout.
// AI Prompts: Generate a C# utility that strips whitespace‑only text nodes and normalizes attribute ordering before DOM comparison. | Enhance the CompareElements method to ignore case differences, optional HTML attributes, and empty style attributes. | Show how to configure HtmlSaveOptions to output external CSS files instead of inline styles for easier DOM validation.

using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using Aspose.Cells;

namespace AsposeCellsHtmlDomValidation
{
    // Creates an Excel workbook, saves it as HTML5 using Aspose.Cells HtmlSaveOptions, loads the output with XmlDocument, defines the target XML structure, and recursively compares both DOM trees to confirm that the HTML table layout, cell IDs and values match the specification. Ideal for automated regression testing of Excel‑to‑HTML conversions.
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create a workbook and populate it ----------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Header");
                sheet.Cells["A2"].PutValue("Row 1");
                sheet.Cells["B2"].PutValue(123);
                sheet.Cells["A3"].PutValue("Row 2");
                sheet.Cells["B3"].PutValue(456);

                // ---------- Configure HTML save options ----------
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Save as HTML5 for modern browsers
                    HtmlVersion = HtmlVersion.Html5,
                    // Ensure JavaScript compatibility for browsers without JS support
                    IsJsBrowserCompatible = true,
                    // Do not embed images as base64 to keep DOM simple
                    ExportImagesAsBase64 = false
                };

                // ---------- Save workbook as HTML ----------
                string htmlPath = "GeneratedOutput.html";
                workbook.Save(htmlPath, htmlOptions);

                // ---------- Load generated HTML into an XML DOM ----------
                if (!File.Exists(htmlPath))
                {
                    Console.WriteLine("Generated HTML file not found.");
                    return;
                }

                XmlDocument generatedDoc = new XmlDocument();
                generatedDoc.Load(htmlPath);

                // ---------- Load expected XML layout ----------
                string expectedXmlContent = @"
                    <html>
                        <head></head>
                        <body>
                            <table>
                                <tr><td id='A1'>Header</td></tr>
                                <tr><td id='A2'>Row 1</td><td id='B2'>123</td></tr>
                                <tr><td id='A3'>Row 2</td><td id='B3'>456</td></tr>
                            </table>
                        </body>
                    </html>";
                XmlDocument expectedDoc = new XmlDocument();
                expectedDoc.LoadXml(expectedXmlContent);

                // ---------- Compare DOM structures ----------
                if (generatedDoc.DocumentElement == null || expectedDoc.DocumentElement == null)
                {
                    Console.WriteLine("One of the XML documents is empty.");
                    return;
                }

                bool isMatch = CompareElements(generatedDoc.DocumentElement, expectedDoc.DocumentElement);
                Console.WriteLine(isMatch
                    ? "Generated HTML matches the expected layout."
                    : "Generated HTML does NOT match the expected layout.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Recursively compare two XmlElements.
        static bool CompareElements(XmlElement htmlElem, XmlElement xmlElem)
        {
            // Element name must match (case‑insensitive).
            if (!string.Equals(htmlElem.Name, xmlElem.Name, StringComparison.OrdinalIgnoreCase))
                return false;

            // Compare attributes (ignoring order).
            if (htmlElem.Attributes.Count != xmlElem.Attributes.Count)
                return false;

            foreach (XmlAttribute htmlAttr in htmlElem.Attributes)
            {
                XmlAttribute xmlAttr = xmlElem.Attributes[htmlAttr.Name];
                if (xmlAttr == null || xmlAttr.Value != htmlAttr.Value)
                    return false;
            }

            // Compare child elements count (excluding insignificant text nodes).
            List<XmlElement> htmlChildren = GetSignificantChildElements(htmlElem);
            List<XmlElement> xmlChildren = GetSignificantChildElements(xmlElem);

            if (htmlChildren.Count != xmlChildren.Count)
                return false;

            for (int i = 0; i < htmlChildren.Count; i++)
            {
                if (!CompareElements(htmlChildren[i], xmlChildren[i]))
                    return false;
            }

            return true;
        }

        // Helper to get element children, ignoring whitespace text nodes.
        static List<XmlElement> GetSignificantChildElements(XmlElement element)
        {
            var list = new List<XmlElement>();
            foreach (XmlNode child in element.ChildNodes)
            {
                if (child.NodeType == XmlNodeType.Element)
                {
                    list.Add((XmlElement)child);
                }
                else if (child.NodeType == XmlNodeType.Text && !string.IsNullOrWhiteSpace(child.Value))
                {
                    // Wrap text nodes in a temporary element for comparison.
                    XmlDocument tempDoc = new XmlDocument();
                    XmlElement wrapper = tempDoc.CreateElement("text");
                    wrapper.InnerText = child.Value.Trim();
                    list.Add(wrapper);
                }
            }
            return list;
        }
    }
}
