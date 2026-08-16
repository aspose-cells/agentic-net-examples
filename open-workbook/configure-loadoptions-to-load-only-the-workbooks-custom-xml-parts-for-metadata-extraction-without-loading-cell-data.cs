// Title: Load custom XML parts only with Aspose.Cells LoadOptions (C#) – skip cell data
// Description: Demonstrates how to configure LoadOptions with a Structure filter and KeepUnparsedData disabled so that only the workbook's custom XML parts are loaded. The example opens the file, enumerates the CustomXmlParts collection, converts each part to a UTF‑8 string, and prints the XML without loading any worksheet cells, delivering fast metadata extraction.
// Keywords: Aspose.Cells LoadOptions C# | custom XML parts Excel | load workbook structure only | skip cell data Aspose | .NET Excel metadata extraction | performance optimization Aspose.Cells
// Common Searches: Aspose.Cells load only custom XML parts | C# load workbook structure without cells | extract Excel custom XML metadata using Aspose | how to avoid loading cell data in Aspose.Cells | read custom XML parts from .xlsx in .NET
// Developer Intent: Open an Excel file just to read its custom XML parts while preventing any cell data from being loaded.
// Use Cases: Indexing document metadata for search engines without the overhead of full workbook parsing. | Transforming embedded XML schemas into JSON or other formats for integration pipelines. | Validating the presence and structure of custom XML parts in automated quality checks.
// AI Prompts: Write C# code that uses Aspose.Cells LoadOptions with a Structure filter to read only custom XML parts from an Excel workbook. | Explain the impact of setting KeepUnparsedData to false when loading a workbook for metadata extraction. | Show how to iterate through the CustomXmlParts collection and output each part as a UTF‑8 string after loading the workbook without cell data.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace LoadCustomXmlOnly
{
    // Demonstrates how to configure LoadOptions with a Structure filter and KeepUnparsedData disabled so that only the workbook's custom XML parts are loaded. The example opens the file, enumerates the CustomXmlParts collection, converts each part to a UTF‑8 string, and prints the XML without loading any worksheet cells, delivering fast metadata extraction.
    class Program
    {
        static void Main()
        {
            // Path to the source workbook that contains custom XML parts
            string sourcePath = "WorkbookWithCustomXml.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Error: The file \"{sourcePath}\" was not found.");
                return;
            }

            try
            {
                // Load only the workbook structure (no cell data) to improve performance
                LoadOptions loadOptions = new LoadOptions
                {
                    LoadFilter = new LoadFilter(LoadDataFilterOptions.Structure),
                    KeepUnparsedData = false
                };

                // Load the workbook with the specified options
                Workbook workbook = new Workbook(sourcePath, loadOptions);

                // Access the collection of custom XML parts
                var customXmlParts = workbook.CustomXmlParts;

                // Output information about the loaded custom XML parts
                Console.WriteLine($"Number of custom XML parts loaded: {customXmlParts.Count}");
                for (int i = 0; i < customXmlParts.Count; i++)
                {
                    // Retrieve the XML data as a string for demonstration
                    string xmlData = Encoding.UTF8.GetString(customXmlParts[i].Data);
                    Console.WriteLine($"--- Custom XML Part {i + 1} ---");
                    Console.WriteLine(xmlData);
                }
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }

            // No need to save the workbook since we only needed metadata extraction
        }
    }
}
