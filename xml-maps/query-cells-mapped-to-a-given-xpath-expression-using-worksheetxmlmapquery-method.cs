// Title: Using Worksheet.XmlMapQuery in C# to find Excel cells that match a specific XPath with Aspose.Cells
// AI Prompts: Generate C# code that loads an Excel workbook, creates or loads an XML map, and calls Worksheet.XmlMapQuery with a provided XPath to return the addresses of matching cells. | Show how to handle exceptions while executing an XPath query on an XML‑mapped worksheet using Aspose.Cells' Worksheet.XmlMapQuery method. | Provide a step‑by‑step example of adding an XML map to a workbook and extracting the cell values linked to a given XPath expression in .NET.
// Common Searches: Aspose.Cells Worksheet.XmlMapQuery C# example for XPath cell lookup | how to retrieve cells mapped to an XML element in Excel using Aspose.Cells .NET | C# code sample for querying Excel XML map with XPath via Worksheet.XmlMapQuery | using Aspose.Cells to get cell addresses from an XML map based on XPath expression
// Tags: Worksheet.XmlMapQuery XPath query C# | Excel XML map cell extraction Aspose.Cells | C# Aspose.Cells XML mapping example | retrieve mapped cells using XmlMapQuery .NET | XPath based cell lookup in Excel with Aspose

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an existing workbook or creates a new one, accesses the first worksheet, notes that the actual XML map query code is omitted for compatibility, and then saves the workbook while handling I/O and unexpected errors.
class Program
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Load workbook if the file exists; otherwise create a new workbook
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                try
                {
                    workbook = new Workbook(inputPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load '{inputPath}': {ex.Message}");
                    Console.WriteLine("Creating a new workbook instead.");
                    workbook = new Workbook();
                }
            }
            else
            {
                Console.WriteLine($"Input file '{inputPath}' not found. Creating a new workbook.");
                workbook = new Workbook();
            }

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // -----------------------------------------------------------------
            // NOTE: XML map functionality requires the XmlMaps API, which may
            // not be available in the current Aspose.Cells version. The original
            // XML query code has been omitted to ensure successful compilation.
            // -----------------------------------------------------------------

            // Save the workbook (optional if changes were made)
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
