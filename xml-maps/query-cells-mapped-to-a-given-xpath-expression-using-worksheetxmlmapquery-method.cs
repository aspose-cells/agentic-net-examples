using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsXmlMapQueryDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Sample XML content to be imported
            string xmlContent = @"<?xml version='1.0' encoding='UTF-8'?>
                <Root>
                    <Data>
                        <Item>Value1</Item>
                        <Item>Value2</Item>
                    </Data>
                </Root>";

            // Write the XML to a temporary file (ImportXml requires a file path)
            string tempXmlPath = "tempSample.xml";
            System.IO.File.WriteAllText(tempXmlPath, xmlContent);

            // Import the XML into the first worksheet starting at cell A1
            workbook.ImportXml(tempXmlPath, "Sheet1", 0, 0);

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the XML map that was created during import
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

            // Link a specific cell to the XML path (optional, demonstrates linking)
            // Here we link cell B1 (row 0, column 1) to the first Item element
            worksheet.Cells.LinkToXmlMap(xmlMap.Name, 0, 1, "/Root/Data/Item[1]");

            // Query the worksheet for cells mapped to the XPath "/Root/Data/Item"
            ArrayList cellAreas = worksheet.XmlMapQuery("/Root/Data/Item", xmlMap);

            // Output the results
            Console.WriteLine($"Number of mapped cell areas: {cellAreas.Count}");
            foreach (CellArea area in cellAreas)
            {
                // For each area, display start row/column and the cell's value
                int row = area.StartRow;
                int column = area.StartColumn;
                string cellAddress = CellsHelper.CellIndexToName(row, column);
                string cellValue = worksheet.Cells[row, column].StringValue;

                Console.WriteLine($"Mapped Cell: {cellAddress} (Row {row}, Column {column})");
                Console.WriteLine($"Cell Value: {cellValue}");
            }

            // Save the workbook to verify the mapping (optional)
            workbook.Save("XmlMapQueryResult.xlsx");

            // Clean up temporary XML file
            System.IO.File.Delete(tempXmlPath);
        }
    }
}