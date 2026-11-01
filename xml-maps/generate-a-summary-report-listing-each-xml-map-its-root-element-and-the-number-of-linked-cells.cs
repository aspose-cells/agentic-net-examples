// Title: Generate an Excel worksheet that lists each XML map’s name, root element, and linked‑cell count using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells to load a workbook, add a sheet named "XML Map Summary", and populate it with the XML map name, root element, and the count of linked cells for every map. | Update the program to safely obtain the linked‑cells count via reflection, skip maps where the property is unavailable, and log those map names. | Extend the summary sheet to also record each map’s namespace URI and auto‑fit all columns before saving the workbook.
// Common Searches: how to list all XML maps in an Excel file with Aspose.Cells C# | asp.net count linked cells of XML maps using Aspose.Cells | c# generate report of XML map root element and linked cells count in workbook
// Tags: enumerate xml maps Aspose.Cells C# | retrieve root element name xml map | linked cells count via reflection Aspose.Cells | add summary worksheet xml map details | auto fit columns Aspose.Cells report

using System;
using System.IO;
using Aspose.Cells;

namespace XmlMapSummaryReport
{
    // The program loads an existing workbook, creates a new worksheet called "XML Map Summary", and writes each XML map’s name, root element, and linked‑cell count (using reflection when necessary). It then auto‑fits the columns and saves the workbook as a new file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputFile = "InputWithXmlMaps.xlsx";
                const string outputFile = "Output_With_XMLMap_Summary.xlsx";

                // Verify that the input workbook exists to avoid FileNotFoundException
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Error: The file '{inputFile}' was not found.");
                    return;
                }

                // Load the existing workbook that contains XML maps
                Workbook workbook = new Workbook(inputFile);

                // Add a new worksheet for the summary report
                int reportSheetIndex = workbook.Worksheets.Add();
                Worksheet reportSheet = workbook.Worksheets[reportSheetIndex];
                reportSheet.Name = "XML Map Summary";

                // Write header titles
                reportSheet.Cells[0, 0].PutValue("XML Map Name");
                reportSheet.Cells[0, 1].PutValue("Root Element");
                reportSheet.Cells[0, 2].PutValue("Linked Cells Count");

                // Iterate through each XML map in the workbook
                int rowIndex = 1;
                foreach (XmlMap xmlMap in workbook.Worksheets.XmlMaps)
                {
                    // XML map name
                    string mapName = xmlMap.Name;

                    // Root element name of the XML map
                    string rootElement = xmlMap.RootElementName;

                    // Determine the number of linked cells.
                    // The LinkedCells property may not be available in all versions,
                    // so reflection is used to obtain the count safely.
                    int linkedCellsCount = 0;
                    var linkedCellsProp = xmlMap.GetType().GetProperty("LinkedCells");
                    if (linkedCellsProp != null)
                    {
                        var linkedCellsObj = linkedCellsProp.GetValue(xmlMap, null);
                        var countProp = linkedCellsObj?.GetType().GetProperty("Count");
                        if (countProp != null)
                        {
                            linkedCellsCount = (int)countProp.GetValue(linkedCellsObj, null);
                        }
                    }

                    // Populate the report row
                    reportSheet.Cells[rowIndex, 0].PutValue(mapName);
                    reportSheet.Cells[rowIndex, 1].PutValue(rootElement);
                    reportSheet.Cells[rowIndex, 2].PutValue(linkedCellsCount);

                    rowIndex++;
                }

                // Auto-fit columns for better readability
                reportSheet.AutoFitColumns();

                // Save the workbook with the new summary report
                workbook.Save(outputFile);
                Console.WriteLine($"Summary report saved to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
