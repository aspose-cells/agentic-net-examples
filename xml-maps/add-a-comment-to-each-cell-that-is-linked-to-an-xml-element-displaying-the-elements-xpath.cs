// Title: Add XPath Comments to XML‑Mapped Cells with Aspose.Cells for .NET (C# Example)
// Description: This C# sample creates a workbook, writes a temporary XML file, adds it as an XML map, links selected cells to XML element paths, inserts a comment in each linked cell that shows the XPath, and saves the result as an XLSX file.
// Keywords: Aspose.Cells | C# XML map | add comment to cell | XPath comment | link cell to XML element | Excel workbook | temporary XML file | Aspose.Cells for .NET example | GitHub code snippet | XML mapping documentation
// Common Searches: How to add an XPath comment to a cell linked with an XML map using Aspose.Cells | C# Aspose.Cells example for linking cells to XML and showing path in comments | Add comments to XML‑mapped cells in .NET | Aspose.Cells XML map tutorial with comments | Insert cell comments with XML element paths in Excel via Aspose
// Developer Intent: Insert a comment into each cell that is linked to an XML element, displaying the element’s XPath.
// Use Cases: Document XML‑to‑Excel mappings by annotating each linked cell with its XPath for easier maintenance | Create an audit‑ready spreadsheet that records source XML paths alongside data values | Build a template that visualizes XML structure directly in Excel through cell comments
// AI Prompts: Generate C# code that uses Aspose.Cells to add an XML map, link specific cells, and insert comments containing each cell’s XPath. | Explain step‑by‑step how to create a temporary XML file, add it as an XML map, link cells, and add XPath comments in Aspose.Cells for .NET. | Show how to read, update, or delete the comment text of a cell that is linked to an XML element using Aspose.Cells API.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlCommentDemo
{
    // This C# sample creates a workbook, writes a temporary XML file, adds it as an XML map, links selected cells to XML element paths, inserts a comment in each linked cell that shows the XPath, and saves the result as an XLSX file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Sample XML content
                string xmlContent = @"<Root><Item>Value</Item></Root>";

                // Write XML to a temporary file to be used for the XML map
                string tempXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xml");
                File.WriteAllText(tempXmlPath, xmlContent);

                // Ensure the temporary XML file exists before adding the map
                if (!File.Exists(tempXmlPath))
                    throw new FileNotFoundException("Temporary XML file was not created.", tempXmlPath);

                // Add the XML map using the temporary file path
                int mapIndex = workbook.Worksheets.XmlMaps.Add(tempXmlPath);
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
                xmlMap.Name = "RootMap";

                // Get the first worksheet and its cells collection
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Define cells to link with their corresponding XML element paths
                var links = new (int Row, int Column, string Path)[]
                {
                    (0, 0, "/Root/Item"), // A1
                    (1, 1, "/Root/Item")  // B2
                };

                // Link each cell to the XML map and add a comment showing the XPath
                foreach (var link in links)
                {
                    // Link the cell to the XML map
                    cells.LinkToXmlMap(xmlMap.Name, link.Row, link.Column, link.Path);

                    // Add a comment to the same cell
                    int commentIndex = sheet.Comments.Add(link.Row, link.Column);
                    Comment comment = sheet.Comments[commentIndex];
                    comment.Note = $"Linked to XML path: {link.Path}";
                }

                // Save the workbook
                string outputPath = "XmlLinkedComments.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
