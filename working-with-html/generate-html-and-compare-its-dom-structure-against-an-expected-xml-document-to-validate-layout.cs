// Title: Generate XHTML from an Aspose.Cells workbook and compare its DOM with an expected XML layout using C#
// AI Prompts: Write C# code that creates a workbook with Aspose.Cells, saves it as XHTML via HtmlSaveOptions, and loads the result into an XmlDocument from a MemoryStream. | Implement a recursive C# method that walks two XmlNode trees and determines structural equality while ignoring whitespace‑only text nodes and attribute order. | Add logic to read a reference XML file, compare its DOM to the generated XHTML DOM using the recursive method, and output whether the structures match.
// Common Searches: how to export Aspose.Cells workbook to xhtml and validate the html structure in c# | c# compare generated html dom with expected xml file ignoring whitespace | aspocells htmlsaveoptions xhtml output and xmldocument comparison example | recursive xml node comparison function c# for html validation
// Tags: Aspose.Cells HTMLSaveOptions XHTML export | C# compare XmlDocument DOM structures | ignore whitespace nodes XML comparison C# | load generated HTML from MemoryStream XmlDocument | validate generated HTML against reference layout C#

using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using Aspose.Cells;
using Aspose.Cells.Rendering; // For HtmlSaveOptions

// The sample creates a simple Aspose.Cells workbook, saves it as XHTML using HtmlSaveOptions, loads the XHTML from a MemoryStream into an XmlDocument, reads an expected XML layout file, and recursively compares the two DOM trees while skipping insignificant whitespace, finally reporting whether the structures match.
class Program
{
    static void Main()
    {
        try
        {
            // -------------------------------------------------
            // 1. Create a simple workbook with sample data
            // -------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue("Item1");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["A3"].PutValue("Item2");
            sheet.Cells["B3"].PutValue(200);

            // -------------------------------------------------
            // 2. Save the workbook as XHTML (HTML that is valid XML)
            // -------------------------------------------------
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // The default HtmlVersion is Xhtml; explicit setting is optional and may not be available in older versions.
            // htmlOptions.HtmlVersion = HtmlVersion.Xhtml;

            using (MemoryStream htmlStream = new MemoryStream())
            {
                workbook.Save(htmlStream, htmlOptions);
                htmlStream.Position = 0; // Reset stream for reading

                // -------------------------------------------------
                // 3. Load the generated XHTML into an XmlDocument
                // -------------------------------------------------
                XmlDocument generatedDoc = new XmlDocument();
                generatedDoc.Load(htmlStream);

                // -------------------------------------------------
                // 4. Load the expected XML layout (provided separately)
                // -------------------------------------------------
                const string expectedPath = "expected_layout.xml";
                if (!File.Exists(expectedPath))
                {
                    Console.WriteLine($"Expected layout file not found: {expectedPath}");
                    return;
                }

                XmlDocument expectedDoc = new XmlDocument();
                expectedDoc.Load(expectedPath);

                // -------------------------------------------------
                // 5. Compare the two DOM structures
                // -------------------------------------------------
                XmlNode genRoot = generatedDoc.DocumentElement;
                XmlNode expRoot = expectedDoc.DocumentElement;

                if (genRoot == null || expRoot == null)
                {
                    Console.WriteLine("One of the XML documents does not have a root element.");
                    return;
                }

                bool structuresMatch = CompareNodes(genRoot, expRoot);
                Console.WriteLine("DOM structures match: " + structuresMatch);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }

    // Recursive comparison of two XmlNode trees (ignores whitespace-only text nodes)
    static bool CompareNodes(XmlNode nodeA, XmlNode nodeB)
    {
        if (nodeA == null || nodeB == null)
            return false;

        // Compare node types (Element, Text, etc.)
        if (nodeA.NodeType != nodeB.NodeType)
            return false;

        // Compare element names
        if (nodeA.NodeType == XmlNodeType.Element && nodeA.Name != nodeB.Name)
            return false;

        // Compare attribute collections
        XmlAttributeCollection attrsA = nodeA.Attributes;
        XmlAttributeCollection attrsB = nodeB.Attributes;
        if ((attrsA?.Count ?? 0) != (attrsB?.Count ?? 0))
            return false;

        if (attrsA != null)
        {
            foreach (XmlAttribute attrA in attrsA)
            {
                XmlAttribute? attrB = attrsB?[attrA.Name];
                if (attrB == null || attrA.Value != attrB.Value)
                    return false;
            }
        }

        // Prepare child node lists, skipping insignificant whitespace text nodes
        List<XmlNode> childrenA = GetSignificantChildren(nodeA);
        List<XmlNode> childrenB = GetSignificantChildren(nodeB);
        if (childrenA.Count != childrenB.Count)
            return false;

        // Recursively compare each child pair
        for (int i = 0; i < childrenA.Count; i++)
        {
            if (!CompareNodes(childrenA[i], childrenB[i]))
                return false;
        }

        // For text nodes, compare trimmed values
        if (nodeA.NodeType == XmlNodeType.Text)
        {
            if (nodeA.Value?.Trim() != nodeB.Value?.Trim())
                return false;
        }

        return true;
    }

    // Helper to retrieve child nodes excluding whitespace-only text nodes
    static List<XmlNode> GetSignificantChildren(XmlNode node)
    {
        var list = new List<XmlNode>();
        foreach (XmlNode child in node.ChildNodes)
        {
            if (child.NodeType == XmlNodeType.Text && string.IsNullOrWhiteSpace(child.Value))
                continue; // Skip insignificant whitespace
            list.Add(child);
        }
        return list;
    }
}
