// Title: Retrieve the address of the first cell mapped to the /Invoice/Total XML element using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, accesses its XML map, and returns the address of the first cell mapped to the '/Invoice/Total' element. | Demonstrate how to use reflection to invoke XmlMap.GetMappedCells in Aspose.Cells and obtain the cell address for a specific XML node. | Create a .NET console program that prints the cell name (e.g., C10) linked to the '/Invoice/Total' element in an Excel file using Aspose.Cells.
// Common Searches: aspocells c# get cell address mapped to xml element invoice total | how to find first cell mapped to /Invoice/Total using Aspose.Cells | using reflection to call GetMappedCells in Aspose.Cells .NET | retrieve xml map cell mapping address Aspose.Cells workbook
// Tags: Aspose.Cells GetMappedCells XML map | C# retrieve mapped cell address | reflection invoke XmlMap GetMappedCells | Excel XML map element to cell address | load workbook Aspose.Cells XML map collection

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, accesses its XML map collection (handling version differences via reflection), calls XmlMap.GetMappedCells for the '/Invoice/Total' element, extracts the first CellArea, retrieves the corresponding cell from the first worksheet, and prints the cell's address such as "C10".
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string totalCellAddress = string.Empty;

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"File not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Obtain the XML map collection via reflection (covers different version property names)
                XmlMapCollection xmlMaps = null;
                try
                {
                    PropertyInfo prop = typeof(Workbook).GetProperty("XmlMaps") ??
                                        typeof(Workbook).GetProperty("XmlMapCollection");
                    if (prop != null)
                    {
                        xmlMaps = prop.GetValue(workbook) as XmlMapCollection;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error retrieving XML maps: {ex.Message}");
                }

                if (xmlMaps == null || xmlMaps.Count == 0)
                {
                    Console.WriteLine("No XML maps found in the workbook.");
                    return;
                }

                // Get the first XML map
                XmlMap xmlMap = xmlMaps[0];

                // Retrieve mapped cells for the /Invoice/Total element
                CellArea[] mappedAreas = null;
                try
                {
                    // Preferred overload (available in newer versions)
                    MethodInfo method = typeof(XmlMap).GetMethod(
                        "GetMappedCells",
                        new[] { typeof(string), typeof(CellArea[]).MakeByRefType() });

                    if (method != null)
                    {
                        object[] parameters = { "/Invoice/Total", null };
                        method.Invoke(xmlMap, parameters);
                        mappedAreas = parameters[1] as CellArea[];
                    }
                    else
                    {
                        // Fallback: method returning CellArea[] directly
                        MethodInfo fallback = typeof(XmlMap).GetMethod(
                            "GetMappedCells",
                            new[] { typeof(string) });

                        if (fallback != null)
                        {
                            mappedAreas = fallback.Invoke(xmlMap, new object[] { "/Invoice/Total" }) as CellArea[];
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error retrieving mapped cells: {ex.Message}");
                }

                // Output the address of the first mapped cell, if any
                if (mappedAreas != null && mappedAreas.Length > 0)
                {
                    CellArea firstArea = mappedAreas[0];
                    Worksheet sheet = workbook.Worksheets[0];
                    Cell firstCell = sheet.Cells[firstArea.StartRow, firstArea.StartColumn];
                    totalCellAddress = firstCell.Name; // e.g., "C10"

                    Console.WriteLine($"First cell mapped to /Invoice/Total: {totalCellAddress}");
                }
                else
                {
                    Console.WriteLine("No cells are mapped to the specified XML element.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
