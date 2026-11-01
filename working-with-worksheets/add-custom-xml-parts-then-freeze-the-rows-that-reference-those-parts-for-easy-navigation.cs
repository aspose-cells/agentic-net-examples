// Title: Add custom XML parts to an Excel workbook and freeze the rows that reference them using Aspose.Cells for .NET (C#)
// Common Searches: c# Aspose.Cells add custom XML part and write its ID to a worksheet cell | how to freeze rows up to the last populated row with Aspose.Cells .NET | store GUID identifiers for custom XML parts in Excel using Aspose.Cells | Aspose.Cells freeze panes based on dynamic row count in C# | add multiple custom XML parts to a workbook and navigate with frozen rows
// Tags: custom xml part insertion Aspose.Cells C# | write part identifier to worksheet cells | freeze panes up to max data row Aspose.Cells | generate GUID for xml part Aspose.Cells | save workbook with custom xml parts .xlsx

using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCustomXmlAndFreeze
{
    // The example creates a new workbook, inserts several XML strings as custom XML parts, writes a generated GUID for each part into column A, determines the last row containing data, freezes all rows up to that row, and saves the file as an .xlsx workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Sample XML strings to be added as custom XML parts
                string[] xmlParts = new string[]
                {
                    "<Customer><Id>1</Id><Name>John Doe</Name></Customer>",
                    "<Customer><Id>2</Id><Name>Jane Smith</Name></Customer>",
                    "<Customer><Id>3</Id><Name>Bob Johnson</Name></Customer>"
                };

                // Add each XML string as a custom XML part and write its identifier into the sheet
                for (int i = 0; i < xmlParts.Length; i++)
                {
                    try
                    {
                        // Convert XML string to UTF‑8 byte array as required by the API
                        byte[] xmlData = Encoding.UTF8.GetBytes(xmlParts[i]);

                        // Add the custom XML part (schemaData is optional, pass null)
                        int partIndex = workbook.CustomXmlParts.Add(xmlData, null);

                        // Generate an identifier for the part (GUID). 
                        // If the library version provides CustomXmlPart.Id, you can replace this with that value.
                        string partId = Guid.NewGuid().ToString();

                        // Write the identifier into column A (zero‑based row index)
                        cells[i, 0].PutValue(partId);
                    }
                    catch (Exception exPart)
                    {
                        Console.WriteLine($"Error adding XML part at index {i}: {exPart.Message}");
                    }
                }

                // Determine the last row that contains data (zero‑based index)
                int lastRowWithReference = cells.MaxDataRow;

                // Freeze rows up to the last referenced row.
                // FreezePanes(row, column, totalRows, totalColumns)
                // row/column specify the cell just below/right of the frozen pane.
                sheet.FreezePanes(lastRowWithReference + 1, 0, lastRowWithReference + 1, 0);

                // Save the workbook to a file
                string outputPath = "CustomXmlAndFreeze.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
