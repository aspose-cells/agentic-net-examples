using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExportXmlStepThreeDemo
    {
        public static void Run()
        {
            // Load the workbook that contains the XML map
            Workbook wb = new Workbook("input.xlsx");

            // Verify that at least one XML map exists
            if (wb.Worksheets.XmlMaps.Count > 0)
            {
                // Retrieve the first XML map
                XmlMap xmlMap = wb.Worksheets.XmlMaps[0];

                // Export the XML data linked to this map to an output file
                wb.ExportXml(xmlMap.Name, "exported_output.xml");

                Console.WriteLine("XML data exported successfully to exported_output.xml");
            }
            else
            {
                Console.WriteLine("No XML map found in the workbook.");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExportXmlStepThreeDemo.Run();
        }
    }
}