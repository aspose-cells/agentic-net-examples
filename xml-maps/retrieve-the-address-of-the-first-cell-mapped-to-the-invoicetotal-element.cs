using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

class RetrieveMappedCellAddress
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Sample XML containing the /Invoice/Total element
            string xml = @"<Invoice><Total>123</Total></Invoice>";

            // Import the XML into the worksheet to create an XML map
            workbook.ImportXml(xml, "Sheet1", 0, 0);

            // Retrieve the created XML map using reflection (avoids compile‑time dependency on XmlMaps property)
            object xmlMapsObj = workbook.GetType().GetProperty("XmlMaps")?.GetValue(workbook);
            if (xmlMapsObj is IList xmlMaps && xmlMaps.Count > 0)
            {
                // Cast the first map to XmlMap
                XmlMap xmlMap = xmlMaps[0] as XmlMap;
                if (xmlMap != null)
                {
                    // Query the worksheet for cell areas mapped to the specified XML path
                    ArrayList mappedAreas = worksheet.XmlMapQuery("/Invoice/Total", xmlMap);

                    if (mappedAreas.Count > 0)
                    {
                        // Get the first mapped area
                        CellArea firstArea = (CellArea)mappedAreas[0];

                        // Obtain the first cell in that area
                        Cell firstMappedCell = worksheet.Cells[firstArea.StartRow, firstArea.StartColumn];

                        // Output the address of the first mapped cell
                        Console.WriteLine("First cell mapped to /Invoice/Total: " + firstMappedCell.Name);
                    }
                    else
                    {
                        Console.WriteLine("No cells are mapped to /Invoice/Total.");
                    }
                }
                else
                {
                    Console.WriteLine("Failed to cast XML map.");
                }
            }
            else
            {
                Console.WriteLine("No XML maps were created.");
            }

            // Save the workbook (optional) – ensure the directory exists
            string outputPath = "MappedCellDemo.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine("Failed to save workbook: " + saveEx.Message);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}