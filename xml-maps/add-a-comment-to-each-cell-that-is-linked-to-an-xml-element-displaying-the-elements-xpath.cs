// Title: Add XPath Comments to Cells Linked via XML Map with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, import an XML map, link specific cells to XML elements using XPath, and insert a visible comment that displays each element's XPath before saving the file.
// Keywords: Aspose.Cells XML map comment | C# add comment with XPath | link Excel cell to XML element | visible comment Aspose.Cells | XPath annotation Excel .NET
// Common Searches: Aspose.Cells add comment showing XPath | C# link cell to XML map and display path | how to annotate XML‑mapped cells in Excel | visible XPath comment Aspose.Cells .NET | add XML map comments programmatically
// Developer Intent: Insert a visible comment into each cell linked to an XML element that shows the element's XPath.
// Use Cases: Provide auditors with a clear trace of each cell's source XML path. | Help end‑users understand the underlying XML schema while editing linked data. | Create documentation sheets that pair data values with their corresponding XPath for validation purposes.
// AI Prompts: Write C# code using Aspose.Cells that scans a workbook for all XML‑mapped cells and adds a comment containing each cell's full XPath. | Create a method that accepts a Workbook object and adds visible XPath comments to every linked cell without overwriting existing comments. | Explain how to modify the sample so comments are hidden by default and appear only on hover in Excel.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlCommentDemo
{
    // Demonstrates how to create a workbook, import an XML map, link specific cells to XML elements using XPath, and insert a visible comment that displays each element's XPath before saving the file.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Sample XML to create an XML map
                string xmlContent = @"<Transmittals><Issued_Document>Test</Issued_Document></Transmittals>";

                // Write XML to a temporary file if it does not already exist
                string xmlFilePath = "Transmittals.xml";
                if (!File.Exists(xmlFilePath))
                {
                    File.WriteAllText(xmlFilePath, xmlContent);
                }

                // Add the XML map to the workbook using the file path
                int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlFilePath);
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
                xmlMap.Name = "Transmittals_Map";

                // Define the cells to link and their corresponding XPath
                var links = new (int Row, int Column, string Path)[]
                {
                    (0, 0, "/Transmittals/Issued_Document"),
                    (1, 1, "/Transmittals/Issued_Document") // example of another linked cell
                };

                // Link each cell to the XML map and add a comment showing the XPath
                foreach (var link in links)
                {
                    // Link the cell to the XML map
                    cells.LinkToXmlMap(xmlMap.Name, link.Row, link.Column, link.Path);

                    // Add a comment to the same cell
                    int commentIdx = worksheet.Comments.Add(link.Row, link.Column);
                    Comment comment = worksheet.Comments[commentIdx];
                    comment.Note = $"Linked to XML element: {link.Path}";
                    comment.Author = "XmlLinker";
                    comment.IsVisible = true;
                }

                // Save the workbook
                workbook.Save("XmlLinkedComments.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
