// Title: Aspose.Cells C# – Define a Dynamically Expanding Named Range for an Excel Table
// Description: Shows how to create a workbook, add a ListObject (Excel table), define a named range with the structured reference =MyTable[#All], insert a new column, and verify that the named range automatically reflects the expanded column count before and after the insertion. The workbook is saved as ExpandingNamedRangeDemo.xlsx.
// Keywords: Aspose.Cells | C# | named range | dynamic range | structured reference | ListObject | Excel table | auto‑expand range | add column programmatically | Excel automation | range expansion
// Common Searches: Aspose.Cells create expanding named range | C# structured reference dynamic range Excel | auto update named range after inserting column Aspose.Cells | ListObject named range expands with new columns | how to make a named range grow automatically in Aspose.Cells
// Developer Intent: Create a named range that automatically expands when new columns are added to an Excel table using Aspose.Cells for .NET.
// Use Cases: Keep formulas, charts, or pivot tables in sync with a table that may gain additional columns. | Reference a single range across worksheets so any horizontal growth is captured without manual edits. | Build reusable templates where the data area adapts to added metrics or attributes.
// AI Prompts: Generate C# code with Aspose.Cells that defines a named range using =MyTable[#All] and prints the column count before and after adding a column. | Explain why a structured reference like =MyTable[#All] causes a named range to expand automatically when the underlying ListObject changes. | Provide best‑practice error handling for retrieving an updated range after modifying a table’s structure in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, add a ListObject (Excel table), define a named range with the structured reference =MyTable[#All], insert a new column, and verify that the named range automatically reflects the expanded column count before and after the insertion. The workbook is saved as ExpandingNamedRangeDemo.xlsx.
    public class ExpandingNamedRangeDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate initial data for the table (3 rows, 2 columns)
                cells["A1"].PutValue("ID");
                cells["B1"].PutValue("Value");
                cells["A2"].PutValue(1);
                cells["B2"].PutValue(10);
                cells["A3"].PutValue(2);
                cells["B3"].PutValue(20);
                cells["A4"].PutValue(3);
                cells["B4"].PutValue(30);

                // Add a ListObject (Excel Table) covering the data range A1:B4
                int tableIndex = sheet.ListObjects.Add("A1", "B4", true);
                ListObject table = sheet.ListObjects[tableIndex];
                table.DisplayName = "MyTable";

                // Create a named range that refers to the whole table using a structured reference.
                int nameIdx = workbook.Worksheets.Names.Add("ExpandedRange");
                Name namedRange = workbook.Worksheets.Names[nameIdx];
                namedRange.RefersTo = "=MyTable[#All]";

                // Verify the initial size of the named range
                AsposeRange rangeBefore = namedRange.GetRange();
                Console.WriteLine($"Before adding column: Columns = {rangeBefore.ColumnCount}, Rows = {rangeBefore.RowCount}");

                // Insert a new column to the right side of the table (after column B)
                sheet.Cells.InsertColumn(2); // Column index 2 corresponds to column C

                // Add header and data for the new column
                cells["C1"].PutValue("Extra");
                cells["C2"].PutValue(100);
                cells["C3"].PutValue(200);
                cells["C4"].PutValue(300);

                // Retrieve the named range again; it should now reflect the expanded table
                AsposeRange rangeAfter = namedRange.GetRange();
                Console.WriteLine($"After adding column: Columns = {rangeAfter.ColumnCount}, Rows = {rangeAfter.RowCount}");

                // Save the workbook
                string outputPath = "ExpandingNamedRangeDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExpandingNamedRangeDemo.Run();
        }
    }
}
