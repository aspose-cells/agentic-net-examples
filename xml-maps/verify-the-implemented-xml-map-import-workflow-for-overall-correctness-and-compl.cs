using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

namespace XmlMapImportVerification
{
    class Program
    {
        static void Main()
        {
            // Path to the XML file to be loaded
            string xmlPath = "sample.xml";

            // Configure XML load options to enable mapping
            XmlLoadOptions loadOptions = new XmlLoadOptions();
            loadOptions.IsXmlMap = true;                     // Enable XML mapping
            loadOptions.ContainsMultipleWorksheets = false; // Single worksheet import
            loadOptions.ConvertNumericOrDate = true;        // Convert numeric/date values

            // Load the XML file into a workbook using the specified options
            Workbook workbook = new Workbook(xmlPath, loadOptions);

            // Verify that at least one XML map was created during import
            if (workbook.Worksheets.XmlMaps.Count == 0)
            {
                Console.WriteLine("No XML maps were detected after loading the XML file.");
                return;
            }

            // Access the first XML map
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];
            Console.WriteLine($"XML Map Name: {xmlMap.Name}");
            Console.WriteLine($"Root Element: {xmlMap.RootElementName}");

            // Link a specific cell to an element in the XML map
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells.LinkToXmlMap(xmlMap.Name, 0, 0, $"/{xmlMap.RootElementName}/*"); // Link A1

            // Query the worksheet to find cells linked to a particular XML path
            string queryPath = $"/{xmlMap.RootElementName}/*";
            ArrayList linkedAreas = sheet.XmlMapQuery(queryPath, xmlMap);
            Console.WriteLine($"Number of cell areas linked to path '{queryPath}': {linkedAreas.Count}");

            // If any linked area is found, display its coordinates and value
            if (linkedAreas.Count > 0)
            {
                CellArea area = (CellArea)linkedAreas[0];
                string cellValue = sheet.Cells[area.StartRow, area.StartColumn].StringValue;
                Console.WriteLine($"First linked cell at Row {area.StartRow + 1}, Column {area.StartColumn + 1}, Value: {cellValue}");
            }

            // Export the XML data using the map name to verify round‑trip capability
            string exportPath = "exported.xml";
            workbook.ExportXml(xmlMap.Name, exportPath);
            Console.WriteLine($"Exported XML saved to '{exportPath}'.");

            // Save the workbook to an Excel file to complete the workflow
            string excelOutput = "VerifiedWorkbook.xlsx";
            workbook.Save(excelOutput);
            Console.WriteLine($"Workbook saved to '{excelOutput}'.");
        }
    }
}