// Title: Export Workbook to XHTML with AspNet.Cells and Verify DOM against Expected XML (C#)
// Description: Create an Excel workbook, save it as XHTML using Aspose.Cells HtmlSaveOptions, load the file as an XmlDocument, and recursively compare its DOM to a predefined XML layout to confirm structural parity.
// Keywords: Aspose.Cells | HtmlSaveOptions | XHTML export | C# DOM comparison | XmlDocument | Excel to HTML | layout validation | regression testing | continuous integration | compare XmlNode trees
// Common Searches: how to export Aspose.Cells workbook to XHTML | C# compare generated HTML DOM with expected XML | validate Aspose.Cells HTML output structure | Aspose.Cells HtmlSaveOptions XHtml example | recursive XmlNode comparison C#
// Developer Intent: Ensure the XHTML generated from an Aspose.Cells workbook matches a reference XML structure.
// Use Cases: Generate a single‑file XHTML report from a workbook and confirm its hierarchy against a baseline template. | Add automated regression tests that detect unintended changes in the HTML rendering of Excel sheets. | Validate styling, element order, and attribute values of exported HTML in CI pipelines.
// AI Prompts: Write a C# utility that normalizes attribute ordering before comparing two XmlNode trees for order‑insensitive DOM validation. | Create a unit test that uses CompareXmlNodes to assert equality between a generated XHTML file and an expected XML layout. | Suggest enhancements to ignore namespace prefixes while still verifying element names and attribute values during DOM comparison.

using System;
using System.IO;
using System.Xml;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsHtmlDomValidation
{
    // Create an Excel workbook, save it as XHTML using Aspose.Cells HtmlSaveOptions, load the file as an XmlDocument, and recursively compare its DOM to a predefined XML layout to confirm structural parity.
    class Program
    {
        static void Main()
        {
            // Paths for generated HTML and expected XML files
            string htmlPath = "GeneratedOutput.xhtml";
            string expectedXmlPath = "ExpectedLayout.xml";

            // ---------- Create a workbook and populate sample data ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue("Row 1");
            sheet.Cells["B2"].PutValue(123);
            sheet.Cells["A3"].PutValue("Row 2");
            sheet.Cells["B3"].PutValue(456);

            // ---------- Configure HTML save options ----------
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Use XHTML output so the file can be loaded as XML
                HtmlVersion = HtmlVersion.XHtml,
                // Normal layout renders like Excel
                LayoutMode = HtmlLayoutMode.Normal,
                // Optional: embed images as Base64 to keep a single file
                ExportImagesAsBase64 = true
            };

            // ---------- Save workbook as HTML (XHTML) ----------
            workbook.Save(htmlPath, saveOptions);

            // ---------- Load generated HTML as XmlDocument ----------
            XmlDocument generatedDoc = new XmlDocument();
            generatedDoc.Load(htmlPath);

            // ---------- Load expected XML layout ----------
            if (!File.Exists(expectedXmlPath))
            {
                Console.WriteLine($"Expected XML file not found: {expectedXmlPath}");
                return;
            }
            XmlDocument expectedDoc = new XmlDocument();
            expectedDoc.Load(expectedXmlPath);

            // ---------- Compare the two DOM structures ----------
            bool areEqual = CompareXmlNodes(generatedDoc.DocumentElement, expectedDoc.DocumentElement);
            Console.WriteLine(areEqual
                ? "The generated HTML DOM matches the expected layout."
                : "The generated HTML DOM does NOT match the expected layout.");
        }

        // Recursively compare two XmlNode trees (element name, attributes, child order)
        static bool CompareXmlNodes(XmlNode nodeA, XmlNode nodeB)
        {
            if (nodeA == null || nodeB == null)
                return nodeA == nodeB;

            // Compare node types (ignore whitespace text nodes)
            if (nodeA.NodeType != nodeB.NodeType)
                return false;

            if (nodeA.NodeType == XmlNodeType.Element)
            {
                // Compare element names
                if (!string.Equals(nodeA.Name, nodeB.Name, StringComparison.Ordinal))
                    return false;

                // Compare attributes (order-insensitive)
                XmlAttributeCollection attrsA = nodeA.Attributes;
                XmlAttributeCollection attrsB = nodeB.Attributes;
                if (attrsA.Count != attrsB.Count)
                    return false;
                foreach (XmlAttribute attrA in attrsA)
                {
                    XmlAttribute attrB = attrsB[attrA.Name];
                    if (attrB == null || attrA.Value != attrB.Value)
                        return false;
                }

                // Compare child nodes (skip insignificant whitespace)
                XmlNodeList childrenA = GetSignificantChildren(nodeA);
                XmlNodeList childrenB = GetSignificantChildren(nodeB);
                if (childrenA.Count != childrenB.Count)
                    return false;

                for (int i = 0; i < childrenA.Count; i++)
                {
                    if (!CompareXmlNodes(childrenA[i], childrenB[i]))
                        return false;
                }
            }
            else if (nodeA.NodeType == XmlNodeType.Text || nodeA.NodeType == XmlNodeType.CDATA)
            {
                // Compare text content (trim to ignore formatting whitespace)
                if (!string.Equals(nodeA.Value?.Trim(), nodeB.Value?.Trim(), StringComparison.Ordinal))
                    return false;
            }

            return true;
        }

        // Helper to retrieve child nodes excluding pure whitespace text nodes
        static XmlNodeList GetSignificantChildren(XmlNode node)
        {
            XmlDocument tempDoc = new XmlDocument();
            XmlElement wrapper = tempDoc.CreateElement("wrapper");
            tempDoc.AppendChild(wrapper);
            foreach (XmlNode child in node.ChildNodes)
            {
                if (child.NodeType == XmlNodeType.Text && string.IsNullOrWhiteSpace(child.Value))
                    continue; // skip insignificant whitespace
                wrapper.AppendChild(tempDoc.ImportNode(child, true));
            }
            return wrapper.ChildNodes;
        }
    }
}
