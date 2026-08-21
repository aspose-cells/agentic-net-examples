// Title: Conditional Formatting for XML‑Mapped Cells with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to add an XML map to a workbook, link cells to XML elements, retrieve the linked ranges via XmlMapQuery, and apply a rule that highlights values greater than 100 with a light‑green fill before saving the file.
// Keywords: Aspose.Cells C# XML map | conditional formatting XML linked cells | XmlMapQuery example | highlight cells based on XML data | LinkToXmlMap C# | FormatCondition greater than | Excel automation Aspose
// Common Searches: Aspose.Cells XmlMapQuery conditional formatting example | C# highlight XML‑mapped cells in Excel | How to apply conditional formatting to XML linked cells using Aspose.Cells | Create conditional formatting rule for XML map in .NET
// Developer Intent: Automatically emphasize cells that are bound to XML map elements when their numeric values exceed a defined threshold.
// Use Cases: Compliance reports that flag Issued_Document values above 100. | Dashboards where XML‑driven metrics are visually highlighted based on business limits. | Spreadsheets that update formatting instantly when the underlying XML source changes.
// AI Prompts: Generate code to add multiple conditional formatting rules (e.g., <50 red, >150 blue) to the same XML‑linked range. | Show how to save the workbook to a MemoryStream while preserving all XML map links and formatting. | Explain how to refresh the XML data source and re‑apply the formatting without rebuilding the workbook.

using System;
using System.Collections;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlConditionalFormatting
{
    // Demonstrates how to add an XML map to a workbook, link cells to XML elements, retrieve the linked ranges via XmlMapQuery, and apply a rule that highlights values greater than 100 with a light‑green fill before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Sample XML data to create an XML map
                string xml = @"<?xml version='1.0' encoding='UTF-8'?>
<Transmittals>
    <Issued_Document>150</Issued_Document>
    <Issued_Document>80</Issued_Document>
    <Issued_Document>200</Issued_Document>
</Transmittals>";

                // Write XML content to a temporary file (required because Add(string) expects a file path)
                string tempXmlPath = Path.Combine(Path.GetTempPath(), "TempXmlMap.xml");
                try
                {
                    File.WriteAllText(tempXmlPath, xml);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to write temporary XML file: {ex.Message}");
                    return;
                }

                // Add the XML map to the workbook (using the temporary file)
                int mapIndex = workbook.Worksheets.XmlMaps.Add(tempXmlPath);
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
                xmlMap.Name = "Transmittals_Map";

                // Link three cells to the XML elements (A1, A2, A3)
                cells.LinkToXmlMap(xmlMap.Name, 0, 0, "/Transmittals/Issued_Document[1]");
                cells.LinkToXmlMap(xmlMap.Name, 1, 0, "/Transmittals/Issued_Document[2]");
                cells.LinkToXmlMap(xmlMap.Name, 2, 0, "/Transmittals/Issued_Document[3]");

                // Query the worksheet for all cell areas linked to the Issued_Document path
                ArrayList linkedAreas = sheet.XmlMapQuery("/Transmittals/Issued_Document", xmlMap);

                // Create a conditional formatting collection for the worksheet
                int cfIndex = sheet.ConditionalFormattings.Add();
                FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

                // Add each linked area to the conditional formatting range
                foreach (CellArea area in linkedAreas)
                {
                    fcc.AddArea(area);
                }

                // Add a condition: highlight cells with value greater than 100
                int conditionIdx = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "100", null);
                FormatCondition condition = fcc[conditionIdx];

                // Define the style to apply when the condition is met
                Style style = workbook.CreateStyle();
                style.ForegroundColor = Color.LightGreen;
                style.Pattern = BackgroundType.Solid;
                condition.Style = style;

                // Define output file path
                string outputPath = "XmlConditionalFormattingOutput.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
