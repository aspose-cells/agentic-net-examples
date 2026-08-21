// Title: Import XML into an Excel Table with Aspose.Cells for .NET (C#) – automatic column mapping
// Description: Demonstrates how to create a new Workbook, verify an XML source file, import the XML data into the first worksheet, calculate the used range, add a ListObject that spans the data (including headers), assign a display name and a built‑in style, and save the result as an XLSX file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | ImportXml | Excel table | ListObject | XML to Excel | automatic column mapping | Workbook import XML | table style
// Common Searches: Aspose.Cells import XML to Excel table C# | Create ListObject from XML data using Aspose.Cells | Map XML elements to Excel columns automatically | How to add a styled table after importing XML with Aspose | C# code to import XML into worksheet and create table
// Developer Intent: Load an XML file, convert its elements into a formatted Excel table, and save the workbook programmatically.
// Use Cases: Generate a ready‑to‑filter Excel report from XML feeds. | Migrate legacy XML data into styled tables for downstream analytics. | Validate XML availability before import and produce a reusable XLSX file.
// AI Prompts: Show how to set the table DisplayName based on the XML file name. | Add a total row with sum calculations to the ListObject after import. | Import XML into a specific worksheet other than Sheet1 and apply a custom table style.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsXmlImportDemo
{
    // Demonstrates how to create a new Workbook, verify an XML source file, import the XML data into the first worksheet, calculate the used range, add a ListObject that spans the data (including headers), assign a display name and a built‑in style, and save the result as an XLSX file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook wb = new Workbook();

                // Verify that the XML source file exists before importing
                const string xmlPath = "data.xml";
                if (!File.Exists(xmlPath))
                {
                    Console.WriteLine($"Error: XML file '{xmlPath}' not found.");
                    return;
                }

                // Import XML data into the first worksheet starting at cell A1 (row 0, column 0)
                wb.ImportXml(xmlPath, "Sheet1", 0, 0);

                // Get reference to the first worksheet where the XML was imported
                Worksheet sheet = wb.Worksheets[0];
                Cells cells = sheet.Cells;

                // Determine the used range after import
                int lastRow = cells.MaxDataRow;          // zero‑based index of the last row with data
                int lastColumn = cells.MaxDataColumn;   // zero‑based index of the last column with data

                // Calculate the total rows and columns for the table (counts, not zero‑based indexes)
                int totalRows = lastRow + 1;
                int totalColumns = lastColumn + 1;

                // Add a table (ListObject) that covers the imported data range, including headers
                int tableIndex = sheet.ListObjects.Add(0, 0, totalRows, totalColumns, true);
                ListObject table = sheet.ListObjects[tableIndex];

                // Set a display name for the table
                table.DisplayName = "XmlDataTable";

                // Apply a built‑in table style
                table.TableStyleType = TableStyleType.TableStyleMedium9;

                // Save the workbook with the imported data and the created table
                const string outputPath = "XmlDataImported.xlsx";
                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
