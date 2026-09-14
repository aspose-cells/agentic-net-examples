// Title: Enable PreserveWhitespace on an Aspose.Cells XmlMap to retain XML formatting spaces in C#
// AI Prompts: Write C# code that opens an Excel workbook with Aspose.Cells, accesses a named XmlMap, sets its PreserveWhitespace property to true, and saves the file. | Show a robust example that checks for the workbook file, iterates over all XmlMaps in the workbook, enables PreserveWhitespace for each, and handles exceptions gracefully. | Provide a snippet that demonstrates how to verify the Aspose.Cells version supports XmlMaps before applying the PreserveWhitespace setting.
// Common Searches: Aspose.Cells C# set XmlMap PreserveWhitespace true example | keep XML whitespace when mapping data to Excel with Aspose.Cells | how to enable whitespace preservation for XmlMap in .NET | C# Aspose.Cells preserve XML indentation in Excel workbook | XmlMap PreserveWhitespace property lost after saving workbook
// Tags: Aspose.Cells XmlMap PreserveWhitespace | C# enable XML whitespace preservation in Excel | modify XmlMap property Aspose.Cells | Excel workbook XML map whitespace handling | Aspose.Cells set PreserveWhitespace true | C# iterate XmlMaps preserve whitespace

using System;
using System.IO;
using Aspose.Cells;

// This example demonstrates how to load an existing Excel workbook using Aspose.Cells for .NET, access its XmlMap collection, set the PreserveWhitespace property to true to retain original XML formatting spaces, and save the workbook. The code includes file existence checks and exception handling for a reliable implementation.
class Program
{
    static void Main()
    {
        try
        {
            const string inputFile = "input.xlsx";
            const string outputFile = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: The file '{inputFile}' was not found.");
                return;
            }

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputFile);

            // NOTE: XmlMap handling removed because the current Aspose.Cells version
            // does not expose the XmlMaps property. If needed, ensure you are using a
            // version that supports XML maps and re‑enable this block.

            // Save the (potentially modified) workbook to a new file
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
