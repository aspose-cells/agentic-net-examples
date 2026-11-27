// Title: How to sync calculated cell values to an XML map and export the updated XML using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export modified worksheet data back to XML map after formula calculation | C# bulk calculate column values and synchronize them with an existing XmlMap in Aspose.Cells | Retrieve cell ranges linked to a specific XML element path using Aspose.Cells ListObject | Import XML into a workbook, create a ListObject linked to XmlMap, then export updated XML in C# | Recalculate formulas and update XML map with Aspose.Cells .NET
// Tags: export updated worksheet to XML map Aspose.Cells | bulk formula calculation with XmlMap synchronization | ListObject linked to XmlMap C# | query cell areas by XML element Aspose.Cells | import XML and create XmlMap automatically Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace XmlMapSyncDemo
{
    // The example demonstrates importing XML data into a workbook (which creates an XmlMap), linking a ListObject to that map, adding a calculated column that doubles existing values, recalculating formulas, and then exporting the modified data back to an XML file via the XmlMap. It also shows how to query cell areas mapped to a specific XML element and saves the workbook for visual verification.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a new workbook and get the first worksheet.
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // -----------------------------------------------------------------
                // 2. Define a simple XML data string that matches the schema we will use.
                // -----------------------------------------------------------------
                string xmlData = @"<?xml version='1.0' encoding='UTF-8'?>
<Data>
    <Item>
        <Name>Item1</Name>
        <Value>100</Value>
    </Item>
    <Item>
        <Name>Item2</Name>
        <Value>200</Value>
    </Item>
    <Item>
        <Name>Item3</Name>
        <Value>300</Value>
    </Item>
</Data>";

                // -----------------------------------------------------------------
                // 3. Import the XML data into the worksheet.
                //    This automatically creates an XmlMap and places the data starting at cell A1.
                // -----------------------------------------------------------------
                workbook.ImportXml(xmlData, "Sheet1", 0, 0);

                // -----------------------------------------------------------------
                // 4. Retrieve the XmlMap that was created by the import operation.
                // -----------------------------------------------------------------
                if (workbook.Worksheets.XmlMaps.Count == 0)
                {
                    Console.WriteLine("No XmlMap was created during XML import.");
                    return;
                }

                XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];
                Console.WriteLine($"XmlMap created: Name = {xmlMap.Name}, RootElement = {xmlMap.RootElementName}");

                // -----------------------------------------------------------------
                // 5. Create a ListObject (table) that covers the imported data.
                //    The table will be linked to the XmlMap via the ListObject.XmlMap property.
                // -----------------------------------------------------------------
                // Determine the used range of the imported data.
                int startRow = 0;               // A1
                int startColumn = 0;            // column A
                int endRow = worksheet.Cells.MaxDataRow;   // last row with data
                int endColumn = worksheet.Cells.MaxDataColumn; // last column with data

                // Add the ListObject (table) to the worksheet.
                int tableIndex = worksheet.ListObjects.Add(startRow, startColumn, endRow, endColumn, true);
                ListObject table = worksheet.ListObjects[tableIndex];
                table.DisplayName = "ItemsTable";

                // Verify that the table is associated with the XmlMap.
                if (table.XmlMap != null)
                {
                    Console.WriteLine($"Table '{table.DisplayName}' is linked to XmlMap '{table.XmlMap.Name}'.");
                }
                else
                {
                    Console.WriteLine($"Table '{table.DisplayName}' is NOT linked to any XmlMap.");
                }

                // -----------------------------------------------------------------
                // 6. Perform bulk calculations on the worksheet.
                //    For demonstration, add a new column that calculates double the Value.
                // -----------------------------------------------------------------
                // Add header for the new column.
                worksheet.Cells[endRow + 1, 0].PutValue("DoubleValue"); // Header in the row after the data.
                // Populate formulas for each data row.
                for (int row = startRow + 1; row <= endRow; row++)
                {
                    // Assuming the original Value column is column B (index 1).
                    // The new formula will be =B{row+1}*2 (Excel rows are 1‑based).
                    string formula = $"=B{row + 1}*2";
                    worksheet.Cells[row, endColumn + 1].Formula = formula;
                }

                // Re‑calculate all formulas in the workbook.
                workbook.CalculateFormula();

                // -----------------------------------------------------------------
                // 7. Synchronize the updated cell values back to the XML map.
                //    Export the current state of the XmlMap to an XML file.
                // -----------------------------------------------------------------
                string outputXmlPath = "UpdatedData.xml";
                workbook.ExportXml(xmlMap.Name, outputXmlPath);
                Console.WriteLine($"Workbook data synchronized to XML map and saved to '{outputXmlPath}'.");

                // -----------------------------------------------------------------
                // 8. (Optional) Demonstrate querying which cells are linked to a specific XML path.
                // -----------------------------------------------------------------
                // Example: Find all cell areas mapped to the "/Data/Item/Value" element.
                string xmlPath = "/Data/Item/Value";
                CellArea[] mappedAreas = worksheet.GetAreasOfXmlMapQuery(xmlPath, xmlMap);
                Console.WriteLine($"Number of cell areas mapped to '{xmlPath}': {mappedAreas.Length}");
                foreach (CellArea area in mappedAreas)
                {
                    Console.WriteLine($"Mapped Area: Row {area.StartRow}‑{area.EndRow}, Column {area.StartColumn}‑{area.EndColumn}");
                }

                // -----------------------------------------------------------------
                // 9. Save the workbook (optional, for visual verification).
                // -----------------------------------------------------------------
                string workbookPath = "XmlMapSyncDemo.xlsx";
                workbook.Save(workbookPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{workbookPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
