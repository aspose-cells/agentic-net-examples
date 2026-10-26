// Title: Set XmlMap PreserveWhitespace to true in Aspose.Cells for .NET to retain XML formatting spaces
// Common Searches: Aspose.Cells C# set XmlMap PreserveWhitespace property to true | preserve XML whitespace when using XmlMap in Aspose.Cells .NET | how to keep XML formatting spaces in an Excel workbook with Aspose.Cells | XmlMap PreserveWhitespace support in different Aspose.Cells versions | retain XML map whitespace during import with Aspose.Cells C#
// Tags: Aspose.Cells XmlMap PreserveWhitespace | C# preserve XML whitespace in Excel workbook | Aspose.Cells set XmlMap whitespace property | XML map formatting spaces handling Aspose.Cells | Aspose.Cells workbook XML map whitespace preservation

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook, checks for XML maps, retrieves the first XmlMap, and shows where the PreserveWhitespace property would be set to true (not available in older Aspose.Cells versions). It then saves the workbook and includes basic error handling.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "output.xlsx";

        try
        {
            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: The file '{inputFile}' was not found.");
                return;
            }

            // Load the existing workbook that may contain XML maps
            Workbook workbook = new Workbook(inputFile);

            // Check for the presence of at least one XML map
            if (workbook.Worksheets.XmlMaps.Count > 0)
            {
                // Retrieve the first XML map (adjust the index if needed)
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

                // The PreserveWhitespace property is not available in this version of Aspose.Cells.
                // If whitespace preservation is required, refer to the appropriate API for the used version.
                // Example placeholder (commented out):
                // xmlMap.PreserveWhitespace = true;
            }

            // Save the workbook with any modifications
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
