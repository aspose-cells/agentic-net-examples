// Title: Find a cell containing a specific XML attribute value in a mapped Excel worksheet using Aspose.Cells Worksheet.Cells.Find (C#)
// AI Prompts: Use Aspose.Cells FindOptions to perform a case‑insensitive search for the XML attribute value "123" in a workbook that has an XML map, and return the address of the matching cell. | Write C# code that verifies an Excel file exists, loads it with Aspose.Cells, and calls Worksheet.Cells.Find to locate the cell that holds a given XML attribute string.
// Common Searches: asp.net locate mapped data cell using aspose.cells | c# find cell by mapped xml value in worksheet | how to get cell address for a specific value in an XML‑mapped Excel file with Aspose.Cells | search for a particular value in a workbook that has an XML map using C#
// Tags: worksheet.cells.find xml attribute lookup | aspose.cells findoptions for xml mapped cells | c# locate xml mapped cell in excel | aspose.cells retrieve cell address from xml map | excel file existence validation with aspose.cells

using Aspose.Cells;
using System;
using System.IO;

// The example checks that 'MappedData.xlsx' exists, loads it with Aspose.Cells, accesses the first worksheet, and uses Worksheet.Cells.Find with a FindOptions object (searching cell values, case‑insensitive) to locate the cell containing the XML attribute value "123". If found, it prints the cell name and its string value.
class Program
{
    static void Main()
    {
        const string filePath = "MappedData.xlsx";

        // Ensure the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook that already has XML data mapped to cells
            Workbook workbook = new Workbook(filePath);
            Worksheet worksheet = workbook.Worksheets[0];

            // The XML attribute value we want to locate in the worksheet
            string xmlAttributeValue = "123";

            // Configure find options: search cell values, case‑insensitive
            FindOptions findOptions = new FindOptions
            {
                LookInType = LookInType.Values,
                CaseSensitive = false
                // LookAtType = LookAtType.WholeContent // Uncomment if the enum is available in your version
                // SearchOrder and SearchDirection are omitted for compatibility with older API versions
            };

            // Locate the cell that contains the XML attribute value
            Cell foundCell = worksheet.Cells.Find(xmlAttributeValue, null, findOptions);

            if (foundCell != null)
            {
                Console.WriteLine($"Cell found at {foundCell.Name} with value '{foundCell.StringValue}'.");
            }
            else
            {
                Console.WriteLine("No cell containing the specified XML attribute value was found.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
