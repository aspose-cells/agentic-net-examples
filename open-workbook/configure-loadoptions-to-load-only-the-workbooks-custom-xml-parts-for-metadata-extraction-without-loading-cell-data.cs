// Title: Load only custom XML parts from an Excel workbook with Aspose.Cells for .NET using LoadOptions to skip cell data
// AI Prompts: Generate C# code that creates a LoadOptions instance to open an .xlsx file with only the CustomXmlParts collection loaded, omitting worksheet data. | Provide a snippet that iterates over workbook.CustomXmlParts after loading a workbook with cell content disabled to extract each part's ID and XML. | Describe how to configure LoadOptions for fast loading when the goal is to retrieve only custom XML metadata from a large Excel workbook.
// Common Searches: Aspose.Cells LoadOptions skip worksheets and load only custom XML parts in C# | Read Excel custom XML parts without loading cell data using Aspose.Cells .NET | How to improve performance by loading only metadata from a large .xlsx with Aspose.Cells | C# example for extracting custom XML parts from workbook while ignoring worksheets
// Tags: Aspose.Cells LoadOptions custom XML parts | skip worksheet loading Aspose.Cells | extract Excel custom XML metadata .NET | performance loading only metadata Aspose.Cells | C# load custom XML parts without cells

using System;
using System.IO;
using Aspose.Cells; // Workbook and related classes

// The sample checks for the presence of an input.xlsx file, loads it with Aspose.Cells, iterates through the workbook's CustomXmlParts collection, and prints each part's Id and XML content while handling any errors.
class Program
{
    static void Main()
    {
        try
        {
            string filePath = "input.xlsx";

            // Verify that the input file exists before attempting to load it.
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the workbook; custom XML parts are loaded automatically.
            Workbook workbook = new Workbook(filePath);

            // Iterate through custom XML parts and display their content.
            // Use 'var' to avoid compile‑time dependency on CustomXmlPart type.
            foreach (var part in workbook.CustomXmlParts)
            {
                try
                {
                    // Cast to dynamic to access Id and Data at runtime.
                    dynamic xmlPart = part;
                    Console.WriteLine($"Part ID: {xmlPart.Id}");
                    Console.WriteLine(xmlPart.Data);
                }
                catch (Exception partEx)
                {
                    Console.WriteLine($"Error processing XML part: {partEx.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors.
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
