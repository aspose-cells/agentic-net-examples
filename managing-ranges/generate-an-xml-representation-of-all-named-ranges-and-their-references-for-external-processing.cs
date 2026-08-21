// Title: Export All Named Ranges (including external links) to XML with Aspose.Cells for .NET
// Description: C# example that loads an Excel workbook, enumerates every defined name, extracts its ReferredArea objects (handling external links), builds a well‑formed XML document with escaped attributes for name, sheet, cell or range coordinates, and saves it as NamedRanges.xml.
// Keywords: Aspose.Cells | .NET | C# | named ranges | XML export | GetReferredAreas | external link | Workbook Names | range address | CellsHelper | code sample
// Common Searches: Aspose.Cells export named ranges to XML | C# list defined names with external references | Generate XML of workbook named ranges | GetReferredAreas example Aspose.Cells | Save named range addresses as XML in .NET
// Developer Intent: Create an XML file that lists every named range in a workbook together with its sheet, cell or area coordinates and any external workbook reference.
// Use Cases: Integrate the XML output into a metadata‑driven reporting pipeline. | Audit workbooks for external links and compliance by parsing the generated file. | Synchronize named range definitions across multiple files by importing the XML data.
// AI Prompts: Write a parser that reads NamedRanges.xml and returns a dictionary of range names to their addresses or external file references. | Extend the code to add a "SourceWorkbook" attribute on the <NamedRanges> root element. | Add comprehensive error handling that skips missing sheets, logs problematic names, and continues processing.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Alias to avoid conflict with System.Range
    using AsposeRange = Aspose.Cells.Range;

    // C# example that loads an Excel workbook, enumerates every defined name, extracts its ReferredArea objects (handling external links), builds a well‑formed XML document with escaped attributes for name, sheet, cell or range coordinates, and saves it as NamedRanges.xml.
    class NamedRangesToXml
    {
        public static void Run()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "NamedRanges.xml";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook
                Workbook wb = new Workbook(inputPath);

                // Prepare XML output
                StringBuilder xml = new StringBuilder();
                xml.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                xml.AppendLine("<NamedRanges>");

                // Iterate through all defined names in the workbook
                foreach (Name name in wb.Worksheets.Names)
                {
                    // Escape the name to be XML‑safe
                    string escapedName = System.Security.SecurityElement.Escape(name.Text);
                    xml.AppendLine($"  <NamedRange Name=\"{escapedName}\">");

                    // Get all referred areas (including external links)
                    ReferredArea[] areas = name.GetReferredAreas(true);
                    if (areas != null && areas.Length > 0)
                    {
                        foreach (ReferredArea area in areas)
                        {
                            xml.Append("    <Reference");

                            // External link information
                            if (area.IsExternalLink)
                            {
                                string extFile = System.Security.SecurityElement.Escape(area.ExternalFileName);
                                xml.Append($" ExternalFile=\"{extFile}\"");
                            }

                            // Sheet name
                            string sheet = System.Security.SecurityElement.Escape(area.SheetName);
                            xml.Append($" Sheet=\"{sheet}\"");

                            // Area or single cell
                            if (area.IsArea)
                            {
                                string start = CellsHelper.CellIndexToName(area.StartRow, area.StartColumn);
                                string end = CellsHelper.CellIndexToName(area.EndRow, area.EndColumn);
                                xml.Append($" Start=\"{start}\" End=\"{end}\"");
                            }
                            else
                            {
                                string cell = CellsHelper.CellIndexToName(area.StartRow, area.StartColumn);
                                xml.Append($" Cell=\"{cell}\"");
                            }

                            xml.AppendLine(" />");
                        }
                    }
                    else
                    {
                        // Fallback: try to get a simple range if no ReferredArea objects are returned
                        AsposeRange range = name.GetRange();
                        if (range != null)
                        {
                            string address = System.Security.SecurityElement.Escape(range.Address);
                            xml.AppendLine($"    <Reference Address=\"{address}\" />");
                        }
                    }

                    xml.AppendLine("  </NamedRange>");
                }

                xml.AppendLine("</NamedRanges>");

                // Save the generated XML to a file
                File.WriteAllText(outputPath, xml.ToString());

                Console.WriteLine($"Named ranges exported to \"{outputPath}\"");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            NamedRangesToXml.Run();
        }
    }
}
