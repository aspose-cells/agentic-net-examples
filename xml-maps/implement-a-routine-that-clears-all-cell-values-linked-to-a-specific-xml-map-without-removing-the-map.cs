// Title: Clear values of cells bound to a specific XML map in an Excel workbook with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# clear data from cells linked to an XML map without deleting the map | remove XML map bound cell values in an .xlsx file using Aspose.Cells API | C# iterate worksheets and clear XML map data source cells with Aspose.Cells | GetDataSourceCells to clear XML map values in Aspose.Cells .NET | preserve XML map definition while clearing its cell values in Aspose.Cells
// Tags: clear XML map bound cells Aspose.Cells | Aspose.Cells GetDataSourceCells usage | preserve XML map definition .NET | iterate worksheets clear XML map values | C# Aspose.Cells XML map data removal

using System;
using System.IO;
using Aspose.Cells;

// Loads a workbook, locates the named XML map, retrieves each worksheet's data source cell areas via GetDataSourceCells, clears the values in those cells while keeping the map definition intact, and saves the updated file. Includes error handling for missing files, absent maps, and Aspose.Cells versions lacking XML map support.
public class XmlMapHelper
{
    /// <param name="inputFile">Path to the source workbook.</param>
    /// <param name="xmlMapName">Name of the XML map whose linked cells should be cleared.</param>
    /// <param name="outputFile">Path where the modified workbook will be saved.</param>
    public static void ClearXmlMapValues(string inputFile, string xmlMapName, string outputFile)
    {
        // Verify input file exists
        if (!File.Exists(inputFile))
            throw new FileNotFoundException("Input workbook not found.", inputFile);

        // Load the workbook
        Workbook workbook = new Workbook(inputFile);

        try
        {
            // Use dynamic to access XmlMaps and related members (handles versions where they are not strongly typed)
            dynamic wbDyn = workbook;
            dynamic xmlMap = wbDyn.XmlMaps[xmlMapName];

            if (xmlMap == null)
                throw new ArgumentException($"XML map '{xmlMapName}' not found in the workbook.");

            // Iterate through each worksheet and clear cells bound to the XML map
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Get the data source areas for this worksheet
                CellArea[] areas = xmlMap.GetDataSourceCells(ws.Name);
                if (areas == null || areas.Length == 0)
                    continue;

                foreach (CellArea area in areas)
                {
                    for (int row = area.StartRow; row <= area.EndRow; row++)
                    {
                        for (int col = area.StartColumn; col <= area.EndColumn; col++)
                        {
                            // Clear the cell value
                            ws.Cells[row, col].PutValue(null);
                        }
                    }
                }
            }

            // Ensure the output directory exists
            string outDir = Path.GetDirectoryName(outputFile);
            if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
                Directory.CreateDirectory(outDir);

            // Save the modified workbook
            workbook.Save(outputFile);
        }
        catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException ex)
        {
            // Thrown when the loaded Aspose.Cells version does not support XML maps
            throw new NotSupportedException("The current Aspose.Cells version does not support XML map operations.", ex);
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            string inputPath = "input.xlsx";
            string mapName = "MyXmlMap";
            string outputPath = "output.xlsx";

            XmlMapHelper.ClearXmlMapValues(inputPath, mapName, outputPath);
            Console.WriteLine("XML map values cleared and workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
