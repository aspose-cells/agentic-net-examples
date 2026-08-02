// Title: Load only custom XML parts from an Excel workbook using Aspose.Cells LoadOptions (C#)
// Description: Demonstrates how to configure Aspose.Cells LoadOptions with a Structure LoadFilter and KeepUnparsedData disabled so the workbook opens without cell data, allowing fast extraction of embedded custom XML parts for metadata processing.
// Keywords: Aspose.Cells | LoadOptions | LoadFilter.Structure | custom XML parts | Excel metadata extraction | C# | skip cell data | performance optimization
// Common Searches: Aspose.Cells load only custom XML parts | Read Excel custom XML without loading worksheets | LoadOptions Structure option C# | Extract embedded XML from Excel file | How to avoid loading cell data in Aspose.Cells
// Developer Intent: Open an Excel file solely to access its custom XML parts while preventing worksheet cell data from being loaded into memory.
// Use Cases: Index embedded XML metadata from large workbooks without the overhead of loading sheet contents. | Validate or transform custom XML schemas in Excel files while keeping memory usage minimal. | Batch‑process many spreadsheets to generate a catalog of their custom XML parts.
// AI Prompts: Generate C# code that uses Aspose.Cells LoadOptions with LoadFilter.Structure to read only custom XML parts from an Excel workbook. | Explain the impact of the KeepUnparsedData flag on loading custom XML parts and recommend the optimal setting for metadata extraction. | Show how to stream each custom XML part directly to a file instead of converting it to a string.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Loading;

namespace AsposeCellsCustomXmlLoad
{
    // Demonstrates how to configure Aspose.Cells LoadOptions with a Structure LoadFilter and KeepUnparsedData disabled so the workbook opens without cell data, allowing fast extraction of embedded custom XML parts for metadata processing.
    class Program
    {
        static void Main()
        {
            // Path to the workbook that contains custom XML parts
            string inputPath = "WorkbookWithCustomXml.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            try
            {
                // Create LoadOptions instance and configure it to load only the workbook structure
                LoadOptions loadOptions = new LoadOptions
                {
                    LoadFilter = new LoadFilter(LoadDataFilterOptions.Structure),
                    KeepUnparsedData = false
                };

                // Load the workbook using the configured options
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Access and enumerate the custom XML parts
                Console.WriteLine($"Number of custom XML parts: {workbook.CustomXmlParts.Count}");
                for (int i = 0; i < workbook.CustomXmlParts.Count; i++)
                {
                    // Retrieve the XML data as a string (UTF-8 encoding assumed)
                    byte[] xmlData = workbook.CustomXmlParts[i].Data;
                    string xmlString = System.Text.Encoding.UTF8.GetString(xmlData);
                    Console.WriteLine($"Custom XML Part {i + 1} content:");
                    Console.WriteLine(xmlString);
                    Console.WriteLine(new string('-', 40));
                }
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
