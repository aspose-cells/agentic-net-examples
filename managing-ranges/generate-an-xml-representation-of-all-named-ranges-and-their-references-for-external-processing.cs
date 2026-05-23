using System;
using System.Xml;
using Aspose.Cells;

namespace AsposeCellsNamedRangeExport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input Excel file path (change as needed)
            string inputPath = "input.xlsx";
            // Output XML file path
            string outputPath = "NamedRanges.xml";

            // Load the workbook (uses the provided load rule)
            Workbook workbook = new Workbook(inputPath);

            // Prepare XML writer with indentation for readability
            XmlWriterSettings settings = new XmlWriterSettings
            {
                Indent = true,
                IndentChars = "  "
            };

            using (XmlWriter writer = XmlWriter.Create(outputPath, settings))
            {
                // Start the XML document
                writer.WriteStartDocument();
                writer.WriteStartElement("NamedRanges");

                // Iterate through all defined names in the workbook
                foreach (Name definedName in workbook.Worksheets.Names)
                {
                    writer.WriteStartElement("NamedRange");
                    // Use the name's text as the identifier
                    writer.WriteAttributeString("Name", definedName.Text);

                    // Retrieve all referred areas (including external links) – uses GetReferredAreas rule
                    ReferredArea[] referredAreas = definedName.GetReferredAreas(true);

                    if (referredAreas != null)
                    {
                        foreach (ReferredArea area in referredAreas)
                        {
                            writer.WriteStartElement("Reference");
                            writer.WriteAttributeString("IsExternal", area.IsExternalLink.ToString());

                            if (area.IsExternalLink)
                            {
                                writer.WriteAttributeString("ExternalFile", area.ExternalFileName);
                            }

                            writer.WriteAttributeString("Sheet", area.SheetName);
                            writer.WriteAttributeString("IsArea", area.IsArea.ToString());

                            // Convert cell indices to A1 style addresses
                            string startAddress = CellsHelper.CellIndexToName(area.StartRow, area.StartColumn);
                            writer.WriteAttributeString("Start", startAddress);

                            if (area.IsArea)
                            {
                                string endAddress = CellsHelper.CellIndexToName(area.EndRow, area.EndColumn);
                                writer.WriteAttributeString("End", endAddress);
                            }

                            writer.WriteEndElement(); // Reference
                        }
                    }

                    writer.WriteEndElement(); // NamedRange
                }

                writer.WriteEndElement(); // NamedRanges
                writer.WriteEndDocument();
            }

            Console.WriteLine($"Named ranges exported to XML file: {outputPath}");
        }
    }
}