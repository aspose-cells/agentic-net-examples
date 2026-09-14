// Title: Enumerate all XML maps in an Excel workbook and log each map’s name and root element with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel file using Aspose.Cells, accesses the Workbook.Worksheets.XmlMaps collection, and writes each XmlMap’s Name and RootElementName to the console. | Provide a snippet that demonstrates how to iterate over XmlMap objects in a workbook and output their details for debugging in a .NET application.
// Common Searches: asp.net list xml maps in an Excel file using Aspose.Cells | c# get xml map name and root element from workbook Aspose.Cells | how to debug xml maps in Excel with Aspose.Cells .NET | enumerate XmlMapCollection in Aspose.Cells C# example | retrieve xml map details from workbook for logging Aspose.Cells
// Tags: list xml maps Aspose.Cells | log xml map name root element C# | debug excel xml maps .NET | retrieve workbook xml map details | iterate XmlMapCollection Aspose.Cells

using System;
using Aspose.Cells;

// Loads an Excel workbook with Aspose.Cells, accesses its Worksheets.XmlMaps collection, and writes each XML map’s Name and RootElementName to the console for debugging.
class Program
{
    static void Main()
    {
        // Load the workbook (replace the path with your actual file)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the collection of XML maps associated with the workbook
        XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;

        // Enumerate each XML map and output its name and root element name
        foreach (XmlMap map in xmlMaps)
        {
            Console.WriteLine($"Map Name: {map.Name}, Root Element: {map.RootElementName}");
        }
    }
}
