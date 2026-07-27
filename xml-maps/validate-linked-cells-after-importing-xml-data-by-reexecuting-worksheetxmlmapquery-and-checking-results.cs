using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsXmlMapValidation
{
    public class ValidateLinkedCells
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Sample XML data to import
                string xmlData = @"<?xml version='1.0' encoding='UTF-8'?>
<Root>
    <Data>
        <Item>Value1</Item>
        <Item>Value2</Item>
    </Data>
</Root>";

                // Import the XML data into the first worksheet starting at cell A1
                // This creates an XML map automatically
                workbook.ImportXml(xmlData, "Sheet1", 0, 0);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the automatically created XML map
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

                // OPTIONAL: Manually link a cell to a specific XML path
                // This demonstrates linking after import
                worksheet.Cells.LinkToXmlMap(xmlMap.Name, 5, 0, "/Root/Data/Item");

                // Define the XML path we want to validate
                string queryPath = "/Root/Data/Item";

                // Re‑execute the XmlMapQuery to obtain all cell areas linked to the path
                ArrayList linkedAreas = worksheet.XmlMapQuery(queryPath, xmlMap);

                // Check the query result
                if (linkedAreas.Count == 0)
                {
                    Console.WriteLine($"No cells are linked to the path '{queryPath}'.");
                }
                else
                {
                    Console.WriteLine($"Found {linkedAreas.Count} linked cell area(s) for path '{queryPath}':");

                    // Iterate through each CellArea and display its address and value
                    foreach (CellArea area in linkedAreas)
                    {
                        // For simplicity, assume each area is a single cell (StartRow/StartColumn)
                        string cellName = CellsHelper.CellIndexToName(area.StartRow, area.StartColumn);
                        string cellValue = worksheet.Cells[area.StartRow, area.StartColumn].StringValue;

                        Console.WriteLine($"- Cell {cellName}: \"{cellValue}\"");
                    }
                }

                // Save the workbook (optional, demonstrates lifecycle rule)
                string outputPath = "ValidatedXmlMap.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ValidateLinkedCells.Run();
        }
    }
}