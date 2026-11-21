// Title: Load only custom XML parts from an Excel workbook using Aspose.Cells LoadOptions in C# (skip worksheet data)
// Common Searches: C# Aspose.Cells load workbook without loading worksheets | How to read custom XML parts from an XLSX file with Aspose.Cells | Extract Excel custom XML parts without loading cell data in .NET | LoadOptions to skip sheet data when opening an Excel file using Aspose.Cells | Retrieve workbook metadata only with Aspose.Cells in C#
// Tags: Aspose.Cells LoadOptions custom XML extraction | C# load workbook without worksheet data | extract custom XML parts from XLSX | metadata‑only workbook loading Aspose.Cells | skip cell data using LoadOptions

using Aspose.Cells;
using System;
using System.IO;

// Demonstrates how to create a LoadOptions object for an XLSX file, open the workbook with Aspose.Cells, and iterate the CustomXmlParts collection to obtain each part's ID and XML content. The example notes that the current Aspose.Cells API does not provide a dedicated flag to load only custom XML parts, so the workbook is opened normally while cell data is effectively ignored for metadata extraction.
class Program
{
    static void Main()
    {
        // Path to the source workbook
        string sourcePath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Error: The file \"{sourcePath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook (no special option for loading only custom XML parts in current API)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Extract and display custom XML parts
            if (workbook.CustomXmlParts.Count == 0)
            {
                Console.WriteLine("No custom XML parts found in the workbook.");
            }
            else
            {
                foreach (object obj in workbook.CustomXmlParts)
                {
                    // Use dynamic to access properties of the custom XML part at runtime
                    dynamic part = obj;
                    Console.WriteLine($"Part ID: {part.Id}");
                    Console.WriteLine($"XML Content: {part.Data}");
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
