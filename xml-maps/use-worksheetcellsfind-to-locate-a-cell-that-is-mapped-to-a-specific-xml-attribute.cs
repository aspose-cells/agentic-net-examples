// Title: Locate a cell containing a specific XML map attribute (e.g., CustomerID) using Worksheet.Cells.Find in Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook, uses Worksheet.Cells.Find with FindOptions.LookInType.Values to locate the cell holding the XML map attribute 'CustomerID', and prints the cell address. | Generate a C# example that searches the first worksheet for any exact XML attribute name supplied at runtime, using Aspose.Cells FindOptions, and returns the corresponding cell reference.
// Common Searches: how to use Worksheet.Cells.Find to locate an XML map attribute in an Excel file with Aspose.Cells C# | Aspose.Cells FindOptions LookInType.Values example for searching XML attribute names | retrieve cell address of a specific XML mapped column header using Aspose.Cells | search for exact XML attribute value in worksheet cells with Aspose.Cells .NET | C# code to find the cell that contains a given XML attribute name in an .xlsx workbook
// Tags: Worksheet.Cells.Find XML attribute lookup | FindOptions LookInType.Values Aspose.Cells | search XML mapped attribute cell C# | locate XML attribute in .xlsx using Aspose.Cells | Aspose.Cells cell address retrieval for XML map

using Aspose.Cells;
using System;
using System.IO;

// The example loads 'input.xlsx', searches the first worksheet for the exact XML map attribute name (e.g., CustomerID) using Worksheet.Cells.Find with LookInType.Values, prints the found cell's address, and saves the workbook as 'output.xlsx'.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // The XML attribute name we want to locate in the worksheet
            string xmlAttributeName = "CustomerID";

            // Configure find options:
            // - LookInType.Values searches cell values (not formulas or notes)
            // - Default LookAtType matches the whole cell content; no explicit setting needed
            FindOptions findOptions = new FindOptions
            {
                LookInType = LookInType.Values
            };

            // Perform the search starting from the first cell (null startCell)
            Cell foundCell = worksheet.Cells.Find(xmlAttributeName, null, findOptions);

            if (foundCell != null)
            {
                Console.WriteLine($"Attribute '{xmlAttributeName}' found at cell {foundCell.Name}");
            }
            else
            {
                Console.WriteLine($"Attribute '{xmlAttributeName}' not found in the worksheet.");
            }

            // Save the workbook (optional if modifications were made)
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
