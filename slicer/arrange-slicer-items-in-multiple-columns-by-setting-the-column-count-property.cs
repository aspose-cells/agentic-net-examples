// Title: Create an Excel ListObject table from a range and save as .xlsx using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# example to add a ListObject table to a worksheet from a cell range | how to set a display name for an Excel table using Aspose.Cells in .NET | C# code to populate worksheet cells and export workbook to XLSX with Aspose.Cells | saving a workbook programmatically with Aspose.Cells for .NET | create table with headers in Aspose.Cells C# tutorial
// Tags: add listobject table Aspose.Cells C# | set table display name Aspose.Cells | populate worksheet cells programmatically C# | save workbook as xlsx Aspose.Cells | create excel table from range Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Tables; // Required for ListObject class

namespace SlicerColumnExample
{
    // The sample creates a new Workbook, renames the first worksheet to "Data", fills cells A1:B7 with sample data, adds a ListObject (Excel table) covering that range with headers, assigns the display name "SampleTable" to the table, and saves the workbook as "SlicerMultipleColumns.xlsx" using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet and rename it
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Data";

                // Populate sample data
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["A5"].PutValue("D");
                sheet.Cells["B5"].PutValue(40);
                sheet.Cells["A6"].PutValue("E");
                sheet.Cells["B6"].PutValue(50);
                sheet.Cells["A7"].PutValue("F");
                sheet.Cells["B7"].PutValue(60);

                // Create a table (list object) from the data range A1:B7
                // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
                int tableIndex = sheet.ListObjects.Add(0, 0, 7, 2, true);
                ListObject table = sheet.ListObjects[tableIndex];

                // Set a display name for the table (Name property is not available in this version)
                table.DisplayName = "SampleTable";

                // Save the workbook
                workbook.Save("SlicerMultipleColumns.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
