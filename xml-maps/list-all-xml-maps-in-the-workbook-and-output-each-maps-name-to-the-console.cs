// Title: How to list all XML maps in an Excel workbook and print each map’s name using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# enumerate XML maps in an Excel file | Get names of XML maps from a workbook with Aspose.Cells .NET | C# code to display all XML map names in an .xlsx using Aspose.Cells | Iterate through XmlMapCollection in Aspose.Cells example | How to check for XML maps in an Excel workbook with Aspose.Cells
// Tags: enumerate xml maps Aspose.Cells | xml map collection iteration C# | list workbook xml maps .NET | retrieve xml map names Aspose.Cells | excel xml map enumeration Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook (defaulting to input.xlsx or a path supplied via command line), accesses the Worksheets.XmlMaps collection, reports the count of XML maps, and writes each map's Name property to the console, with handling for missing files and runtime exceptions.
class Program
{
    static void Main(string[] args)
    {
        // Determine the workbook file to process.
        // If a path is supplied as a command‑line argument it is used,
        // otherwise a default file name "input.xlsx" in the current directory is assumed.
        string workbookPath = args.Length > 0 ? args[0] : "input.xlsx";

        // Verify that the file exists before attempting to load it.
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: The file \"{workbookPath}\" does not exist.");
            return;
        }

        try
        {
            // Load the workbook.
            Workbook workbook = new Workbook(workbookPath);

            // Access the collection of XML maps attached to the workbook.
            XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;

            // Report the result.
            if (xmlMaps.Count == 0)
            {
                Console.WriteLine("No XML maps were found in the workbook.");
            }
            else
            {
                Console.WriteLine($"XML maps found: {xmlMaps.Count}");
                for (int i = 0; i < xmlMaps.Count; i++)
                {
                    XmlMap map = xmlMaps[i];
                    // The Name property holds the map's identifier.
                    Console.WriteLine($"- Map {i + 1}: {map.Name}");
                }
            }
        }
        catch (Exception ex)
        {
            // Any exception (e.g., corrupted file, unsupported format) is reported.
            Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
        }
    }
}
