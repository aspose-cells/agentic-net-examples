using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main()
        {
            LoadXmlSchemaDemo.Run();
        }
    }

    public class LoadXmlSchemaDemo
    {
        public static void Run()
        {
            // Create a new, empty workbook
            Workbook workbook = new Workbook();

            // Add the XML schema (XSD) file to the workbook's XML maps collection.
            int mapIndex = workbook.Worksheets.XmlMaps.Add("schema.xsd");

            // Retrieve the added XmlMap (optional, useful for further configuration)
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];

            // Assign a friendly name to the map (optional)
            xmlMap.Name = "MySchemaMap";

            // Save the workbook; the XML map is now embedded in the file.
            workbook.Save("WorkbookWithXmlMap.xlsx");
        }
    }
}