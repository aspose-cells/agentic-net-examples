// Title: Generate a CSV report of XML map names, root elements, and linked cell counts with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# list all XML maps in a workbook and export their details to CSV | How to count cells linked to an XML map across worksheets using Aspose.Cells | C# retrieve XML map root element name and linked cell total with Aspose.Cells | Create a summary report of Excel XML maps using Aspose.Cells .NET library | Export XML map information (name, root element, cell count) to a CSV file in C#
// Tags: Aspose.Cells XML map enumeration | CSV export of XML map details | linked cell counting with XmlMapQuery | root element extraction from Excel XML maps | C# workbook XML map summary

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook with Aspose.Cells, iterates its XmlMapCollection, obtains each map's name and root element, uses XmlMapQuery to count every cell linked to the map across all worksheets, and writes the map name, root element, and total linked cells to a CSV file while also printing the information to the console.
class Program
{
    static void Main(string[] args)
    {
        // Determine the workbook file to process.
        string workbookPath = args.Length > 0 ? args[0] : "input.xlsx";

        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: Workbook file \"{workbookPath}\" does not exist.");
            return;
        }

        try
        {
            // Load the workbook.
            Workbook workbook = new Workbook(workbookPath);

            // Access the collection of XML maps defined for the workbook.
            XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;

            if (xmlMaps.Count == 0)
            {
                Console.WriteLine("No XML maps were found in the workbook.");
                return;
            }

            // Prepare a report file.
            string reportPath = "XmlMapSummaryReport.csv";
            using (StreamWriter reportWriter = new StreamWriter(reportPath))
            {
                // Write CSV header.
                reportWriter.WriteLine("XmlMapName,RootElementName,LinkedCellCount");

                // Iterate through each XML map.
                for (int i = 0; i < xmlMaps.Count; i++)
                {
                    XmlMap xmlMap = xmlMaps[i];
                    string mapName = xmlMap.Name ?? $"Map_{i}";
                    string rootElement = xmlMap.RootElementName ?? string.Empty;

                    // Count linked cells across all worksheets.
                    int linkedCellCount = 0;

                    // The path to query – using the root element name.
                    // The XmlMapQuery method expects an XPath-like string.
                    // Using \"//\" + rootElement will match any occurrence of the root element.
                    string queryPath = string.IsNullOrEmpty(rootElement) ? string.Empty : $"//{rootElement}";

                    // Iterate each worksheet to collect areas linked to this map.
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // XmlMapQuery is obsolete but still functional; it returns an ArrayList of CellArea objects.
                        // If the queryPath is empty, skip the call to avoid exceptions.
                        if (string.IsNullOrEmpty(queryPath))
                            continue;

                        ArrayList areas = sheet.XmlMapQuery(queryPath, xmlMap);
                        foreach (object obj in areas)
                        {
                            if (obj is CellArea area)
                            {
                                int rows = area.EndRow - area.StartRow + 1;
                                int cols = area.EndColumn - area.StartColumn + 1;
                                linkedCellCount += rows * cols;
                            }
                        }
                    }

                    // Output to console.
                    Console.WriteLine($"Map: {mapName}");
                    Console.WriteLine($"  Root Element: {rootElement}");
                    Console.WriteLine($"  Linked Cells : {linkedCellCount}");
                    Console.WriteLine();

                    // Write to CSV.
                    reportWriter.WriteLine($"{EscapeCsv(mapName)},{EscapeCsv(rootElement)},{linkedCellCount}");
                }
            }

            Console.WriteLine($"Summary report written to \"{reportPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }

    // Helper to escape CSV fields that may contain commas or quotes.
    private static string EscapeCsv(string field)
    {
        if (field == null)
            return string.Empty;

        if (field.Contains(",") || field.Contains("\"") || field.Contains("\n"))
        {
            string escaped = field.Replace("\"", "\"\"");
            return $"\"{escaped}\"";
        }
        return field;
    }
}
