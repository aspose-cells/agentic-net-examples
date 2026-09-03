// Title: Generate an Excel summary of XML maps with root elements and linked cell counts using Aspose.Cells for .NET
// AI Prompts: Write a C# console application that opens an existing .xlsx workbook, uses reflection to iterate over its Workbook.XmlMaps collection, extracts each map's Name and RootElementName, counts every cell linked to the map across all worksheets via GetXmlMapLinkedCells, and saves the map name, root element, and linked‑cell total into a new workbook called XmlMapSummaryReport.xlsx. | Enhance the XML‑map summary program to also record the worksheet names that contain linked cells for each map and output the expanded data as a CSV file.
// Common Searches: how to list XML maps in an Excel file using Aspose.Cells C# | C# count cells linked to an XML map in a workbook with Aspose.Cells | Aspose.Cells get root element name of XML map via reflection | generate XML map summary report Excel Aspose.Cells .NET | retrieve XmlMapLinkedCells for each worksheet in Aspose.Cells
// Tags: Aspose.Cells enumerate XmlMaps collection | C# count linked cells for XML map | export XML map details to Excel workbook | reflection access to XmlMaps in Aspose.Cells | GetXmlMapLinkedCells usage example

using System;
using System.IO;
using System.Collections;
using Aspose.Cells;

// The program loads a source .xlsx workbook, uses reflection to obtain the XmlMaps collection, iterates each map to read its Name and RootElementName, counts all cells linked to the map across every worksheet via GetXmlMapLinkedCells, and writes the map name, root element, and linked‑cell count into a new Excel file named XmlMapSummaryReport.xlsx.
class XmlMapSummaryReport
{
    static void Main()
    {
        // Input workbook path
        string sourcePath = "input.xlsx";

        // Verify that the input file exists
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Error: Input file \"{sourcePath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook that may contain XML maps
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Use reflection to obtain the XmlMaps collection (may not exist in older versions)
            var xmlMapsProp = sourceWorkbook.GetType().GetProperty("XmlMaps");
            if (xmlMapsProp == null)
            {
                Console.WriteLine("The loaded workbook does not support XML maps in this Aspose.Cells version.");
                return;
            }

            var xmlMapsObj = xmlMapsProp.GetValue(sourceWorkbook) as IEnumerable;
            if (xmlMapsObj == null)
            {
                Console.WriteLine("No XML maps found in the workbook.");
                return;
            }

            // Create a new workbook for the summary report
            Workbook reportWorkbook = new Workbook();
            Worksheet sheet = reportWorkbook.Worksheets[0];
            sheet.Name = "XML Map Summary";

            // Write header row
            sheet.Cells[0, 0].PutValue("XML Map Name");
            sheet.Cells[0, 1].PutValue("Root Element");
            sheet.Cells[0, 2].PutValue("Linked Cells Count");

            int reportRow = 1; // start writing data from the second row

            // Iterate through each XML map using reflection
            foreach (object xmlMap in xmlMapsObj)
            {
                // Retrieve map name and root element via reflection
                string mapName = xmlMap.GetType().GetProperty("Name")?.GetValue(xmlMap) as string ?? "N/A";
                string rootElement = xmlMap.GetType().GetProperty("RootElementName")?.GetValue(xmlMap) as string ?? "N/A";

                int linkedCellCount = 0;

                // Iterate worksheets to count linked cells for the current map
                foreach (Worksheet ws in sourceWorkbook.Worksheets)
                {
                    // Locate the GetXmlMapLinkedCells method (signature may vary)
                    var method = ws.GetType().GetMethod("GetXmlMapLinkedCells", new Type[] { xmlMap.GetType() });
                    if (method == null)
                    {
                        // Fallback: try overload that accepts a string map name
                        method = ws.GetType().GetMethod("GetXmlMapLinkedCells", new Type[] { typeof(string) });
                        if (method != null)
                        {
                            var areas = method.Invoke(ws, new object[] { mapName }) as Array;
                            linkedCellCount += CountCellsInAreas(areas);
                        }
                        continue;
                    }

                    var linkedAreas = method.Invoke(ws, new object[] { xmlMap }) as Array;
                    linkedCellCount += CountCellsInAreas(linkedAreas);
                }

                // Write data to the report sheet
                sheet.Cells[reportRow, 0].PutValue(mapName);
                sheet.Cells[reportRow, 1].PutValue(rootElement);
                sheet.Cells[reportRow, 2].PutValue(linkedCellCount);
                reportRow++;
            }

            // Auto‑fit columns for readability
            sheet.AutoFitColumns();

            // Output workbook path
            string reportPath = "XmlMapSummaryReport.xlsx";

            // Save the summary report
            reportWorkbook.Save(reportPath);
            Console.WriteLine($"Report saved to \"{reportPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to count cells in an array of CellArea objects (using reflection)
    private static int CountCellsInAreas(Array areas)
    {
        if (areas == null) return 0;
        int count = 0;
        foreach (object area in areas)
        {
            var startRowProp = area.GetType().GetProperty("StartRow");
            var endRowProp = area.GetType().GetProperty("EndRow");
            var startColProp = area.GetType().GetProperty("StartColumn");
            var endColProp = area.GetType().GetProperty("EndColumn");

            if (startRowProp == null || endRowProp == null || startColProp == null || endColProp == null)
                continue;

            int startRow = (int)startRowProp.GetValue(area);
            int endRow = (int)endRowProp.GetValue(area);
            int startCol = (int)startColProp.GetValue(area);
            int endCol = (int)endColProp.GetValue(area);

            int rows = endRow - startRow + 1;
            int cols = endCol - startCol + 1;
            count += rows * cols;
        }
        return count;
    }
}
