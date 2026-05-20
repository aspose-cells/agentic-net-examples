using System;
using System.IO;
using Aspose.Cells;

class ExportXmlWithoutDeclaration
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Id");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");

            // Create an XML map (using a simple schema as a placeholder)
            int mapIndex = workbook.Worksheets.XmlMaps.Add("<Schema><Element><Id/><Name/></Element></Schema>");
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
            xmlMap.Name = "SampleMap";

            // Export the XML using the map name and output file path.
            // The ExportXml method does not require ExportXmlOptions in older versions of Aspose.Cells.
            string outputPath = "output_without_declaration.xml";
            workbook.ExportXml(xmlMap.Name, outputPath);

            // If the XML declaration needs to be removed and the current version does not support it via options,
            // perform a simple post‑processing step.
            if (File.Exists(outputPath))
            {
                string[] lines = File.ReadAllLines(outputPath);
                if (lines.Length > 0 && lines[0].StartsWith("<?xml"))
                {
                    File.WriteAllLines(outputPath, lines[1..]); // Write file without the first line
                }
            }

            Console.WriteLine("XML exported successfully to " + Path.GetFullPath(outputPath));
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}