// Title: C# program to convert an Excel workbook to XHTML with Aspose.Cells and validate the generated HTML DOM against an expected XML layout
// Common Searches: how to export an Excel workbook to XHTML using Aspose.Cells in C# with images embedded as Base64 | C# compare two XML DOM trees while ignoring whitespace and attribute order | validate generated HTML from Aspose.Cells against a reference XML structure | recursive XML node comparison function in .NET | load Excel file, save as HTML with Aspose.Cells and test DOM consistency
// Tags: Aspose.Cells HtmlSaveOptions XHTML export | C# XML DOM structural comparison | ignore whitespace XML node comparison .NET | Excel to HTML conversion with embedded Base64 images | validate generated HTML DOM against reference XML

using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using Aspose.Cells;

namespace AsposeCellsHtmlDomValidation
{
    // The example loads an Excel workbook, saves it as well‑formed XHTML using Aspose.Cells HtmlSaveOptions with images encoded as Base64, then loads both the generated HTML file and a predefined XML file into XmlDocument objects. A recursive routine compares the two DOM trees, ignoring insignificant whitespace and attribute order, and reports whether the HTML layout matches the expected XML structure.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths – adjust as needed or pass via command‑line arguments
            string workbookPath = "input.xlsx";
            string htmlOutputPath = "generated.html";
            string expectedXmlPath = "expected.xml";

            try
            {
                // -------------------------------------------------
                // 1. Load the workbook
                // -------------------------------------------------
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Workbook file not found: {workbookPath}");
                    return;
                }

                Workbook workbook = new Workbook(workbookPath);

                // -------------------------------------------------
                // 2. Save workbook as XHTML (well‑formed XML)
                // -------------------------------------------------
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Keep a single file by embedding images as Base64
                    ExportImagesAsBase64 = true
                };

                workbook.Save(htmlOutputPath, htmlOptions);
                Console.WriteLine($"Workbook saved as HTML to: {htmlOutputPath}");

                // -------------------------------------------------
                // 3. Load generated HTML and expected XML into XmlDocument
                // -------------------------------------------------
                if (!File.Exists(htmlOutputPath))
                {
                    Console.WriteLine($"Generated HTML file not found: {htmlOutputPath}");
                    return;
                }

                XmlDocument generatedDoc = new XmlDocument();
                try
                {
                    generatedDoc.Load(htmlOutputPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load generated HTML as XML: {ex.Message}");
                    return;
                }

                if (!File.Exists(expectedXmlPath))
                {
                    Console.WriteLine($"Expected XML file not found: {expectedXmlPath}");
                    return;
                }

                XmlDocument expectedDoc = new XmlDocument();
                try
                {
                    expectedDoc.Load(expectedXmlPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load expected XML: {ex.Message}");
                    return;
                }

                // -------------------------------------------------
                // 4. Compare DOM structures
                // -------------------------------------------------
                bool areEqual = false;
                if (generatedDoc.DocumentElement != null && expectedDoc.DocumentElement != null)
                {
                    areEqual = CompareXmlNodes(generatedDoc.DocumentElement, expectedDoc.DocumentElement);
                }
                else
                {
                    Console.WriteLine("One of the documents does not contain a root element.");
                }

                Console.WriteLine(areEqual
                    ? "Success: The generated HTML DOM matches the expected XML structure."
                    : "Failure: The generated HTML DOM does NOT match the expected XML structure.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected error occurred:");
                Console.WriteLine(ex.Message);
            }
        }

        // Recursively compare two XML nodes (ignoring insignificant whitespace)
        static bool CompareXmlNodes(XmlNode nodeA, XmlNode nodeB)
        {
            if (nodeA == null && nodeB == null) return true;
            if (nodeA == null || nodeB == null) return false;

            // Compare node types
            if (nodeA.NodeType != nodeB.NodeType) return false;

            // Compare element names
            if (nodeA.NodeType == XmlNodeType.Element && nodeA.Name != nodeB.Name) return false;

            // Compare attributes (order‑independent)
            if (nodeA.Attributes != null || nodeB.Attributes != null)
            {
                int countA = nodeA.Attributes?.Count ?? 0;
                int countB = nodeB.Attributes?.Count ?? 0;
                if (countA != countB) return false;

                foreach (XmlAttribute attrA in nodeA.Attributes)
                {
                    XmlAttribute attrB = nodeB.Attributes[attrA.Name];
                    if (attrB == null || attrA.Value != attrB.Value) return false;
                }
            }

            // Compare text content for text nodes
            if (nodeA.NodeType == XmlNodeType.Text)
            {
                string textA = nodeA.Value?.Trim() ?? string.Empty;
                string textB = nodeB.Value?.Trim() ?? string.Empty;
                if (textA != textB) return false;
            }

            // Compare child nodes (ignoring pure whitespace)
            List<XmlNode> childrenA = GetSignificantChildren(nodeA);
            List<XmlNode> childrenB = GetSignificantChildren(nodeB);

            if (childrenA.Count != childrenB.Count) return false;

            for (int i = 0; i < childrenA.Count; i++)
            {
                if (!CompareXmlNodes(childrenA[i], childrenB[i])) return false;
            }

            return true;
        }

        // Returns child nodes excluding whitespace‑only text nodes
        static List<XmlNode> GetSignificantChildren(XmlNode parent)
        {
            var list = new List<XmlNode>();
            foreach (XmlNode child in parent.ChildNodes)
            {
                if (child.NodeType == XmlNodeType.Whitespace ||
                    child.NodeType == XmlNodeType.SignificantWhitespace)
                {
                    continue; // ignore formatting whitespace
                }

                if (child.NodeType == XmlNodeType.Text && string.IsNullOrWhiteSpace(child.Value))
                {
                    continue; // ignore empty text nodes
                }

                list.Add(child);
            }
            return list;
        }
    }
}
