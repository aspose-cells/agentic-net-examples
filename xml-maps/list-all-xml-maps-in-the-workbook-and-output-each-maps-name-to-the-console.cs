// Title: List XML map names in an Excel workbook with Aspose.Cells for .NET
// Description: Loads a workbook, accesses the Worksheets.XmlMaps collection, checks for presence, and prints each XML map’s Name to the console using the Aspose.Cells C# API.
// Keywords: Aspose.Cells | C# | XML map | list XML maps | Workbook.XmlMaps | enumerate XML maps | Excel XML mapping | retrieve XML map names
// Common Searches: Aspose.Cells get XML map names | C# list XML maps in Excel | enumerate XML maps using Aspose.Cells | retrieve XML map collection .NET | display XML map names from workbook
// Developer Intent: Obtain and display the names of every XML map defined in an Excel file.
// Use Cases: Confirm required XML maps exist before importing XML data. | Log map names for debugging data‑integration pipelines. | Create a quick audit report of XML map definitions in a workbook.
// AI Prompts: Generate C# code with Aspose.Cells that lists all XML map names in a workbook and handles the case when none are present. | Show how to write the retrieved XML map names to a text file instead of the console using Aspose.Cells. | Explain the steps to add a new XML map to a workbook and then enumerate all existing XML maps with Aspose.Cells.

using System;
using Aspose.Cells;

// Loads a workbook, accesses the Worksheets.XmlMaps collection, checks for presence, and prints each XML map’s Name to the console using the Aspose.Cells C# API.
class ListXmlMaps
{
    static void Main()
    {
        // Load an existing workbook (replace "input.xlsx" with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the collection of XML maps in the workbook
        XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;

        // If there are no XML maps, inform the user
        if (xmlMaps.Count == 0)
        {
            Console.WriteLine("No XML maps found in the workbook.");
        }
        else
        {
            // Iterate through each XML map and output its name
            for (int i = 0; i < xmlMaps.Count; i++)
            {
                XmlMap map = xmlMaps[i];
                Console.WriteLine($"XML Map {i}: {map.Name}");
            }
        }
    }
}
