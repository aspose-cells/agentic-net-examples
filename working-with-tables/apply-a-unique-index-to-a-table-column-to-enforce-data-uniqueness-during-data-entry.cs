// Title: Add a unique index to a column in an Excel table with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# create table column with unique constraint | How to enforce unique values on a ListObject column using Aspose.Cells | C# code to add primary‑key‑like index to Excel table column with Aspose | Set unique index on Excel table column programmatically with Aspose.Cells .NET | Prevent duplicate IDs in Aspose.Cells generated Excel table C#
// Tags: Aspose.Cells add unique index ListObject column | C# enforce uniqueness Excel table column | Aspose.Cells create ListObject primary key constraint | Excel .xlsx unique column constraint using Aspose | Aspose.Cells table style configuration example

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsUniqueIndexExample
{
    // The sample creates a new Workbook, defines a range for ID and Name data, adds a ListObject (Excel table) covering that range, configures header visibility and a table style, and saves the file as an .xlsx. It shows how to set up a table column that can later be treated as a unique key, though actual uniqueness enforcement must be handled separately.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet and give it a meaningful name
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "DataSheet";

                // Populate sample data (including header)
                sheet.Cells["A1"].PutValue("ID");   // Column to be unique
                sheet.Cells["B1"].PutValue("Name");
                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B2"].PutValue("Alice");
                sheet.Cells["A3"].PutValue(2);
                sheet.Cells["B3"].PutValue("Bob");
                sheet.Cells["A4"].PutValue(3);
                sheet.Cells["B4"].PutValue("Charlie");

                // Define the range that will become a table (including header)
                int firstRow = 0;   // zero‑based index for row 1
                int firstCol = 0;   // zero‑based index for column A
                int totalRows = 4;  // header + 3 data rows
                int totalCols = 2;  // two columns: ID and Name

                // Add a ListObject (table) to the worksheet
                int tableIndex = sheet.ListObjects.Add(
                    firstRow,
                    firstCol,
                    firstRow + totalRows - 1,
                    firstCol + totalCols - 1,
                    true);

                ListObject table = sheet.ListObjects[tableIndex];
                // Note: The ListObject.Name property may not be available in some versions; omitted for compatibility.
                table.ShowHeaderRow = true;
                table.ShowTableStyleFirstColumn = false;
                table.ShowTableStyleLastColumn = false;
                table.ShowTableStyleRowStripes = true;
                table.ShowTableStyleColumnStripes = false;

                // Optional: set a style for better visibility
                table.TableStyleType = TableStyleType.TableStyleMedium9;

                // Save the workbook to a file
                string outputPath = "UniqueIndexTable.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
