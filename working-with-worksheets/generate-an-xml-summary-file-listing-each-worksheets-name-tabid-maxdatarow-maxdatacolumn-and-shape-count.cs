// Title: C# – Export Worksheet Name, TabId, MaxDataRow, MaxDataColumn & ShapeCount to XML with Aspose.Cells
// Description: A concise C# example that loads an Excel workbook using Aspose.Cells, iterates over every worksheet, captures the sheet name, TabId, zero‑based MaxDataRow, MaxDataColumn, and the number of shapes, then writes these values as attributes of <Worksheet> nodes inside a <WorkbookSummary> XML file.
// Keywords: Aspose.Cells XML export | C# worksheet metadata | Excel TabId Aspose | MaxDataRow MaxDataColumn | shape count worksheet | .NET Excel to XML | worksheet summary generation
// Common Searches: Aspose.Cells generate XML summary of worksheets | C# get TabId, MaxDataRow, MaxDataColumn from Excel | export shape count from Excel sheet using Aspose | write worksheet properties to XML in .NET | how to list all worksheets metadata with Aspose.Cells
// Developer Intent: Create an XML document that lists each worksheet’s name, TabId, last data row, last data column, and shape count using Aspose.Cells for .NET.
// Use Cases: Produce a lightweight documentation file for workbook structure audits. | Feed worksheet boundaries and embedded object counts into a data‑pipeline without opening the Excel file. | Compare two workbook versions by generating XML summaries and diffing the results to spot added/removed sheets, data ranges, or shapes.
// AI Prompts: Generate C# code with Aspose.Cells that writes an XML summary of all worksheets, including Name, TabId, MaxDataRow, MaxDataColumn, and ShapeCount attributes. | Show how to modify XmlWriterSettings to add a custom namespace and XSD schema reference to the WorkbookSummary XML. | Extend the example to include each worksheet’s visibility state (Visible, Hidden, VeryHidden) as an extra XML attribute.

using System;
using System.IO;
using System.Xml;
using Aspose.Cells;

namespace AsposeCellsSummaryDemo
{
    // A concise C# example that loads an Excel workbook using Aspose.Cells, iterates over every worksheet, captures the sheet name, TabId, zero‑based MaxDataRow, MaxDataColumn, and the number of shapes, then writes these values as attributes of <Worksheet> nodes inside a <WorkbookSummary> XML file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file
            string excelPath = "input.xlsx";

            // Path where the XML summary will be saved
            string xmlOutputPath = "WorkbookSummary.xml";

            GenerateWorkbookSummary(excelPath, xmlOutputPath);
        }

        /// <param name="excelPath">Path to the Excel workbook.</param>
        /// <param name="xmlOutputPath">Path to save the generated XML summary.</param>
        static void GenerateWorkbookSummary(string excelPath, string xmlOutputPath)
        {
            // Load the workbook using Aspose.Cells (lifecycle rule: create/load)
            Workbook workbook = new Workbook(excelPath);

            // Prepare an XmlWriter with indentation for readability
            XmlWriterSettings settings = new XmlWriterSettings
            {
                Indent = true,
                IndentChars = "    ",
                NewLineOnAttributes = false
            };

            using (XmlWriter writer = XmlWriter.Create(xmlOutputPath, settings))
            {
                // Start the root element
                writer.WriteStartDocument();
                writer.WriteStartElement("WorkbookSummary");

                // Iterate through all worksheets in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Retrieve required properties
                    string sheetName = sheet.Name;
                    int tabId = sheet.TabId;
                    int maxDataRow = sheet.Cells.MaxDataRow;       // Zero‑based index of the last row containing data
                    int maxDataColumn = sheet.Cells.MaxDataColumn; // Zero‑based index of the last column containing data
                    int shapeCount = sheet.Shapes.Count;           // Number of drawing shapes (pictures, charts, etc.)

                    // Write a <Worksheet> element with attributes for each property
                    writer.WriteStartElement("Worksheet");
                    writer.WriteAttributeString("Name", sheetName);
                    writer.WriteAttributeString("TabId", tabId.ToString());
                    writer.WriteAttributeString("MaxDataRow", maxDataRow.ToString());
                    writer.WriteAttributeString("MaxDataColumn", maxDataColumn.ToString());
                    writer.WriteAttributeString("ShapeCount", shapeCount.ToString());
                    writer.WriteEndElement(); // </Worksheet>
                }

                // Close the root element
                writer.WriteEndElement(); // </WorkbookSummary>
                writer.WriteEndDocument();
            }

            // Optionally, inform the user
            Console.WriteLine($"Workbook summary saved to '{xmlOutputPath}'.");
        }
    }
}
