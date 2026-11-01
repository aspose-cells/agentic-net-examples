// Title: Retrieve the A1 address of the first worksheet cell mapped to the /Invoice/Total XML element using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook, accesses its XML maps, calls GetMappedCells for '/Invoice/Total', and prints the A1 address of the first mapped cell. | Show how to use Aspose.Cells to find the first cell linked to a specific XML element and return its cell name. | Demonstrate extracting the start cell address from the CellArea array returned by GetMappedCells for the '/Invoice/Total' element.
// Common Searches: asp.net aspose.cells get first cell address for xml map element | c# retrieve A1 notation of cell mapped to /Invoice/Total in Excel | how to use GetMappedCells to locate cells linked to an XML element with Aspose.Cells | example of reading XML map cell mapping in Aspose.Cells .NET | find first worksheet cell mapped to an XML schema element using Aspose.Cells
// Tags: Aspose.Cells GetMappedCells first cell address | XML map element to worksheet cell lookup C# | retrieve A1 notation from CellArea Aspose.Cells | Excel XML mapping cell extraction .NET | C# Aspose.Cells XML map cell address

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlMapExample
{
    // The example loads an Excel workbook, accesses its first XML map, uses GetMappedCells to obtain the cell areas mapped to the '/Invoice/Total' element, selects the first area, retrieves the corresponding cell from the first worksheet, and outputs its address in A1 notation.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the input workbook
            string inputPath = "input.xlsx";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            try
            {
                // Load the workbook using dynamic to avoid compile‑time binding issues with XmlMaps
                dynamic workbook = new Workbook(inputPath);

                // Access the XmlMaps collection
                dynamic xmlMaps = workbook.XmlMaps;

                // Ensure that at least one XML map is present
                if (xmlMaps == null || xmlMaps.Count == 0)
                {
                    Console.WriteLine("No XML maps are defined in the workbook.");
                    return;
                }

                // Get the first XML map (assumes the required map is the first one)
                dynamic xmlMap = xmlMaps[0];

                // Retrieve all cell areas mapped to the /Invoice/Total element
                CellArea[] mappedAreas = (CellArea[])xmlMap.GetMappedCells("/Invoice/Total");

                string firstCellAddress = string.Empty;

                if (mappedAreas != null && mappedAreas.Length > 0)
                {
                    // The first CellArea represents the first mapped cell
                    CellArea firstArea = mappedAreas[0];

                    // Access the cell using its start row and column
                    Cell firstCell = workbook.Worksheets[0].Cells[firstArea.StartRow, firstArea.StartColumn];

                    // Get the address in A1 notation (e.g., "B5")
                    firstCellAddress = firstCell.Name;
                }

                if (!string.IsNullOrEmpty(firstCellAddress))
                {
                    Console.WriteLine($"First cell mapped to '/Invoice/Total' is: {firstCellAddress}");
                }
                else
                {
                    Console.WriteLine("No cells are mapped to '/Invoice/Total'.");
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
