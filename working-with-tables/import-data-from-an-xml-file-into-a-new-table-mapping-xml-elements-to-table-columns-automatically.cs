// Title: Create an Excel ListObject from XML file with dynamic column detection using Aspose.Cells for .NET (C#)
// Common Searches: how to convert XML file to Excel table using Aspose.Cells C# | Aspose.Cells import XML and create ListObject with dynamic columns | C# read XML and generate Excel table with automatic headers Aspose | map XML elements to Excel columns programmatically Aspose.Cells | save XML data as styled Excel table using Aspose.Cells .NET
// Tags: Aspose.Cells XML to ListObject conversion | dynamic column detection from XDocument | C# generate Excel table from XML data | apply table style to imported XML worksheet | error handling for missing XML file Aspose.Cells

using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example loads an XML document, automatically derives column names from the first repeating element, writes headers and row data to the first worksheet, creates a ListObject (Excel table) over the populated range with a predefined style, and saves the workbook as an XLSX file, including robust error handling for missing files and load/save failures.
class Program
{
    static void Main()
    {
        // Path to the source XML file
        string xmlPath = "data.xml";

        // Verify that the XML file exists
        if (!File.Exists(xmlPath))
        {
            Console.WriteLine($"Error: XML file not found at path '{xmlPath}'.");
            return;
        }

        XDocument xdoc;
        try
        {
            // Load the XML document
            xdoc = XDocument.Load(xmlPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading XML file: {ex.Message}");
            return;
        }

        XElement? root = xdoc.Root;
        if (root == null)
        {
            Console.WriteLine("The XML file does not contain a root element.");
            return;
        }

        // Identify the first repeating element (e.g., <Item>) to derive column names
        XElement? firstItem = root.Elements().FirstOrDefault();
        if (firstItem == null)
        {
            Console.WriteLine("The XML file does not contain any data elements.");
            return;
        }

        // Extract column names from child element names of the first item
        var columnNames = firstItem.Elements()
                                   .Select(e => e.Name.LocalName)
                                   .ToList();

        if (columnNames.Count == 0)
        {
            Console.WriteLine("No columns were detected in the XML data.");
            return;
        }

        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Write header row with column names
        for (int col = 0; col < columnNames.Count; col++)
        {
            cells[0, col].PutValue(columnNames[col]);
        }

        // Write data rows for each repeating element
        int currentRow = 1;
        foreach (XElement item in root.Elements())
        {
            for (int col = 0; col < columnNames.Count; col++)
            {
                string colName = columnNames[col];
                XElement? element = item.Element(colName);
                string value = element != null ? element.Value : string.Empty;
                cells[currentRow, col].PutValue(value);
            }
            currentRow++;
        }

        // Determine the range of the table (including header)
        int totalRows = currentRow;               // number of rows written (header + data)
        int totalCols = columnNames.Count;        // number of columns

        // Add a ListObject (Excel table) over the populated range
        try
        {
            int tableIndex = sheet.ListObjects.Add(0, 0, totalRows - 1, totalCols - 1, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "XmlDataTable";
            table.TableStyleType = TableStyleType.TableStyleMedium9;
            table.ShowTableStyleFirstColumn = true;
            table.ShowTableStyleLastColumn = true;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error creating Excel table: {ex.Message}");
            // Continue without a table; the data is still written to the sheet.
        }

        // Save the workbook to an XLSX file
        string outputPath = "Output.xlsx";
        try
        {
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
