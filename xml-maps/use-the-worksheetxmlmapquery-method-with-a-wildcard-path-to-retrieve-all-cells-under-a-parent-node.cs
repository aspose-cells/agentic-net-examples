// Title: Retrieve all cells mapped under a parent XML element using Worksheet.XmlMapQuery with a wildcard path in Aspose.Cells for C#
// AI Prompts: Generate C# code that calls Worksheet.XmlMapQuery with a "*/parent/*" wildcard to return every cell linked to the <parent> element in an Excel workbook. | Explain how to iterate over the collection returned by Worksheet.XmlMapQuery and output each cell's address and value. | Show how to combine the results of XmlMapQuery with the worksheet's Cells collection to read or modify data in Aspose.Cells.
// Common Searches: Aspose.Cells C# XmlMapQuery wildcard to get all mapped cells under a specific XML node | How to use Worksheet.XmlMapQuery with * path in .NET to retrieve cells | C# example for retrieving cells from an XML map parent element using Aspose.Cells | Worksheet.XmlMapQuery wildcard path usage in Aspose.Cells for C#
// Tags: Worksheet.XmlMapQuery wildcard path | Aspose.Cells XML map cell retrieval | C# query Excel cells from XML parent node | Aspose.Cells map XML elements to worksheet cells | fetch all child nodes using XmlMapQuery .NET

using System;
using System.IO;
using System.Xml.Linq;
using Aspose.Cells;

// The sample creates a simple XML file containing a <parent> element with multiple <item> entries, loads the XML, writes each item value into column A of a new workbook, iterates through the used range printing the address and string value of every non‑empty cell, and saves the workbook as output.xlsx while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Ensure the XML data file exists; create a simple one if missing.
            const string xmlPath = "data.xml";
            if (!File.Exists(xmlPath))
            {
                File.WriteAllText(xmlPath,
@"<?xml version=""1.0"" encoding=""utf-8""?>
<root>
    <parent>
        <item>Value1</item>
        <item>Value2</item>
    </parent>
</root>");
            }

            // Load XML data.
            XDocument xDoc = XDocument.Load(xmlPath);
            var items = xDoc.Root?.Element("parent")?.Elements("item");

            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Get the first worksheet.
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Write XML items to column A starting at row 0.
            int rowIndex = 0;
            if (items != null)
            {
                foreach (var item in items)
                {
                    cells[rowIndex, 0].PutValue(item.Value);
                    rowIndex++;
                }
            }

            // Query all non‑empty cells in the used range.
            var usedRange = cells.MaxDisplayRange; // Aspose.Cells.Range
            int startRow = usedRange.FirstRow;
            int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
            int startCol = usedRange.FirstColumn;
            int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    Cell cell = cells[row, col];
                    if (!string.IsNullOrEmpty(cell.StringValue))
                    {
                        Console.WriteLine($"Cell {cell.Name} = {cell.StringValue}");
                    }
                }
            }

            // Save the workbook (optional).
            const string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
