// Title: Export All Excel Shapes to XML with Aspose.Cells for .NET
// Description: Loads an Excel workbook, walks through each worksheet and its Shapes collection, and writes key shape attributes—name, type, position, size, visibility, lock state, alternative text, and linked cell—into a well‑structured XML file (ShapesDefinition.xml). This enables external analysis, documentation, and version‑control tracking of shape metadata.
// Keywords: Aspose.Cells | C# | export Excel shapes to XML | shape metadata | worksheet shapes | XML serialization | version control | Excel automation | shape properties | alternative text | linked cell
// Common Searches: export Excel shapes to XML Aspose.Cells | C# extract shape properties from workbook | save shape definitions for version control | list all shapes in Excel file programmatically | Aspose.Cells shape XML output | how to serialize Excel shapes | retrieve alternative text from Excel shapes C#
// Developer Intent: Extract every shape’s metadata from an Excel workbook and serialize it to an XML file for downstream analysis or source‑control comparison.
// Use Cases: Automated CI/CD step that generates ShapesDefinition.xml and diffs it against previous builds to catch unintended layout changes. | Create a catalog of form controls (checkboxes, dropdowns, etc.) by parsing the exported XML for alternative text and linked cell references. | Provide auditors or business analysts with a readable XML report of all visual elements in a spreadsheet. | Enable custom tooling that reads the XML to rebuild or migrate shapes across workbooks.
// AI Prompts: Write code that reads the generated ShapesDefinition.xml and reconstructs the shapes in a new workbook using Aspose.Cells. | Extend the export to include rotation angle, fill color, and line style for each shape. | Create a GitHub Action that runs the export program, commits ShapesDefinition.xml, and fails the build if differences exceed a threshold. | Generate a PowerShell script that executes the compiled exe, captures the XML output, and pushes it to a remote repository.

using System;
using System.IO;
using System.Xml;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeExport
{
    // Loads an Excel workbook, walks through each worksheet and its Shapes collection, and writes key shape attributes—name, type, position, size, visibility, lock state, alternative text, and linked cell—into a well‑structured XML file (ShapesDefinition.xml). This enables external analysis, documentation, and version‑control tracking of shape metadata.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // Prepare the output XML file that will contain all shape definitions
            string outputPath = "ShapesDefinition.xml";

            // Create an XmlWriter with indentation for readability
            XmlWriterSettings settings = new XmlWriterSettings
            {
                Indent = true,
                IndentChars = "  ",
                NewLineOnAttributes = false
            };

            using (XmlWriter writer = XmlWriter.Create(outputPath, settings))
            {
                // Start the root element
                writer.WriteStartDocument();
                writer.WriteStartElement("WorkbookShapes");

                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Write a container element for the current worksheet
                    writer.WriteStartElement("Worksheet");
                    writer.WriteAttributeString("Name", sheet.Name);
                    writer.WriteAttributeString("Index", sheet.Index.ToString());

                    // Iterate through each shape on the worksheet
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Write a shape element with selected properties
                        writer.WriteStartElement("Shape");
                        writer.WriteAttributeString("Name", shape.Name ?? string.Empty);
                        writer.WriteAttributeString("Type", shape.Type.ToString());
                        writer.WriteAttributeString("UpperLeftRow", shape.UpperLeftRow.ToString());
                        writer.WriteAttributeString("UpperLeftColumn", shape.UpperLeftColumn.ToString());
                        writer.WriteAttributeString("LowerRightRow", shape.LowerRightRow.ToString());
                        writer.WriteAttributeString("LowerRightColumn", shape.LowerRightColumn.ToString());
                        writer.WriteAttributeString("Width", shape.Width.ToString());
                        writer.WriteAttributeString("Height", shape.Height.ToString());
                        writer.WriteAttributeString("IsHidden", shape.IsHidden.ToString());
                        writer.WriteAttributeString("IsLocked", shape.IsLocked.ToString());

                        // Optional: include alternative text if present
                        if (!string.IsNullOrEmpty(shape.AlternativeText))
                        {
                            writer.WriteElementString("AlternativeText", shape.AlternativeText);
                        }

                        // Optional: include linked cell information for controls
                        if (!string.IsNullOrEmpty(shape.LinkedCell))
                        {
                            writer.WriteElementString("LinkedCell", shape.LinkedCell);
                        }

                        // Close the Shape element
                        writer.WriteEndElement();
                    }

                    // Close the Worksheet element
                    writer.WriteEndElement();
                }

                // Close the root element
                writer.WriteEndElement();
                writer.WriteEndDocument();
            }

            Console.WriteLine($"All shape definitions have been exported to '{outputPath}'.");
        }
    }
}
