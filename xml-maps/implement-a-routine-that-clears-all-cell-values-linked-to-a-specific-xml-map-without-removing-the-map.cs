// Title: Clear values of cells linked to a specific XML map in an Excel workbook with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# clear data from cells linked to an XML map without deleting the map | GetAreasOfXmlMapQuery example to remove linked cell values in a .NET Excel file | How to preserve an XML map while clearing its linked cell contents using Aspose.Cells | Programmatically clear XML map linked ranges in an Excel workbook with C# | Aspose.Cells sample to clear XML map data but keep the map definition
// Tags: Aspose.Cells clear linked cell ranges | XML map cell area extraction C# | retain XML map after clearing values | Excel workbook modify linked cells Aspose | C# Aspose.Cells clear range preserving formatting

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, finds the specified XmlMap, retrieves all cell areas linked to that map via GetAreasOfXmlMapQuery, clears the values in those ranges with Cells.ClearRange while keeping formatting, and saves the workbook, leaving the XML map definition unchanged.
class Program
{
    /// <param name="workbookPath">Path to the input Excel workbook.</param>
    /// <param name="xmlMapName">Name of the XML map whose linked cells should be cleared.</param>
    /// <param name="outputPath">Path where the modified workbook will be saved.</param>
    static void ClearXmlMapLinkedCells(string workbookPath, string xmlMapName, string outputPath)
    {
        // Verify input file exists
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Input workbook not found: \"{workbookPath}\"");
            return;
        }

        try
        {
            // Load the workbook
            Workbook wb = new Workbook(workbookPath);

            // Locate the requested XmlMap
            XmlMap targetMap = null;
            foreach (XmlMap map in wb.Worksheets.XmlMaps)
            {
                if (string.Equals(map.Name, xmlMapName, StringComparison.OrdinalIgnoreCase))
                {
                    targetMap = map;
                    break;
                }
            }

            if (targetMap == null)
            {
                Console.WriteLine($"XmlMap \"{xmlMapName}\" was not found in the workbook.");
                return;
            }

            Console.WriteLine($"Found XmlMap \"{targetMap.Name}\". Beginning to clear linked cells...");

            // Iterate through worksheets and clear cells linked to the map
            foreach (Worksheet ws in wb.Worksheets)
            {
                // Get all areas linked to the map (empty path returns all linked cells)
                CellArea[] linkedAreas = ws.GetAreasOfXmlMapQuery(string.Empty, targetMap);

                if (linkedAreas == null || linkedAreas.Length == 0)
                    continue; // No linked cells on this sheet

                foreach (CellArea area in linkedAreas)
                {
                    int totalRows = area.EndRow - area.StartRow + 1;
                    int totalColumns = area.EndColumn - area.StartColumn + 1;

                    // Clear values while preserving formatting
                    ws.Cells.ClearRange(area.StartRow, area.StartColumn, totalRows, totalColumns);
                    Console.WriteLine($"Cleared range {area.StartRow},{area.StartColumn} - {area.EndRow},{area.EndColumn} on sheet \"{ws.Name}\".");
                }
            }

            // Ensure output directory exists
            string outDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
                Directory.CreateDirectory(outDir);

            // Save the modified workbook
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\". XML map \"{xmlMapName}\" remains intact.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
        }
    }

    static void Main(string[] args)
    {
        if (args.Length != 3)
        {
            Console.WriteLine("Usage: AsposeCellsClearXmlMap <inputWorkbookPath> <xmlMapName> <outputWorkbookPath>");
            return;
        }

        string inputPath = args[0];
        string mapName   = args[1];
        string outputPath = args[2];

        try
        {
            ClearXmlMapLinkedCells(inputPath, mapName, outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
