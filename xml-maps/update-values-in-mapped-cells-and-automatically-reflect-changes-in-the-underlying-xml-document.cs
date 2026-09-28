// Title: How to update XML‑mapped cells in an Excel workbook and save the changes using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx file with an XML map, changes the values of mapped cells (e.g., Price and Quantity), and saves the workbook to a new file using Aspose.Cells. | Write a step‑by‑step C# example that accesses a worksheet, updates cells linked to XML elements, and persists the workbook with Aspose.Cells.
// Common Searches: c# asp.net update xml mapped cells in existing excel file using aspose.cells | how to edit xml map data in a workbook and save it with Aspose.Cells for .NET | example of changing values of cells linked to XML elements in an xlsx via Aspose.Cells | programmatically modify xml map linked cells and export updated workbook in C#
// Tags: xml map cell update Aspose.Cells | modify mapped cells in xlsx C# | save workbook after xml map edit Aspose.Cells | Aspose.Cells XML map editing example | update Excel cell values linked to XML elements

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing workbook (MappedWorkbook.xlsx) that contains an XML map, updates two mapped cells (B2 and C5) representing XML elements such as <Price> and <Quantity>, and saves the modified workbook as MappedWorkbook_Updated.xlsx using Aspose.Cells for .NET. It notes that the XmlMaps API may not be available in all library versions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputFile = "MappedWorkbook.xlsx";
            const string outputFile = "MappedWorkbook_Updated.xlsx";

            // Verify that the input workbook exists.
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file \"{inputFile}\" not found.");
                return;
            }

            // Load the workbook.
            Workbook workbook = new Workbook(inputFile);

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Update cells that are mapped to XML elements/attributes.
            sheet.Cells["B2"].PutValue(199.99); // Example: mapped to <Price>
            sheet.Cells["C5"].PutValue(42);     // Example: mapped to <Quantity>

            // NOTE: The XmlMaps API may not be available in all Aspose.Cells versions.
            // If needed, XML map export can be implemented using a compatible version of the library.

            // Save the updated workbook.
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved as \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
