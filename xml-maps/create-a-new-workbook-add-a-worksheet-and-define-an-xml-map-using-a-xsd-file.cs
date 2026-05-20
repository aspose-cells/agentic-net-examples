using System;
using Aspose.Cells;

namespace AsposeCellsXmlMapDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (default worksheet is already added)
            Workbook workbook = new Workbook();

            // Add an additional worksheet (optional, demonstrates adding a sheet)
            int newSheetIndex = workbook.Worksheets.Add();
            Worksheet newSheet = workbook.Worksheets[newSheetIndex];
            newSheet.Name = "DataSheet";

            // Path to the XSD file that defines the XML schema
            // Ensure that "schema.xsd" exists at the specified location
            string xsdPath = "schema.xsd";

            // Define an XML map in the workbook using the XSD file
            // The Add method returns the index of the newly added XmlMap
            int xmlMapIndex = workbook.Worksheets.XmlMaps.Add(xsdPath);

            // Optionally set a friendly name for the XML map
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[xmlMapIndex];
            xmlMap.Name = "MyXmlMap";

            // Save the workbook to a file
            workbook.Save("MappedWorkbook.xlsx");

            Console.WriteLine("Workbook created, worksheet added, and XML map defined successfully.");
        }
    }
}