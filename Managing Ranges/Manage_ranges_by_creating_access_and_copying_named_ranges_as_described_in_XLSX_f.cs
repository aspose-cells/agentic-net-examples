using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data in A1:C3
            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // -------------------------------------------------
            // 1. Create a named range "MyData" that refers to A1:C3
            // -------------------------------------------------
            // Create the range object
            AsposeRange sourceRange = cells.CreateRange("A1", "C3");

            // Add the name to the workbook's Names collection
            int nameIndex = workbook.Worksheets.Names.Add("MyData");
            // Set the RefersTo formula to point to the created range
            workbook.Worksheets.Names[nameIndex].RefersTo = $"={sheet.Name}!$A$1:$C$3";

            // -------------------------------------------------
            // 2. Access the named range using Name.GetRange()
            // -------------------------------------------------
            Name named = workbook.Worksheets.Names["MyData"];
            AsposeRange accessedRange = named.GetRange(); // gets the range object

            // Verify the address
            Console.WriteLine($"Accessed Range Address: {accessedRange.RefersTo}");

            // -------------------------------------------------
            // 3. Get all ranges referred by the name (GetRanges)
            // -------------------------------------------------
            AsposeRange[] allRanges = named.GetRanges();
            Console.WriteLine($"Number of ranges referred by 'MyData': {allRanges?.Length ?? 0}");

            // -------------------------------------------------
            // 4. Copy the named range to a new location (E1:G3)
            // -------------------------------------------------
            // Create destination range with same dimensions
            AsposeRange destRange = cells.CreateRange("E1", "G3");
            // Copy data, formulas, formatting etc.
            accessedRange.Copy(destRange);

            // -------------------------------------------------
            // 5. Retrieve the same range using WorksheetCollection.GetRangeByName
            // -------------------------------------------------
            AsposeRange retrievedByCollection = workbook.Worksheets.GetRangeByName("MyData");
            if (retrievedByCollection != null)
            {
                Console.WriteLine($"Retrieved by collection: {retrievedByCollection.RefersTo}");
            }

            // -------------------------------------------------
            // 6. Save the workbook (save rule)
            // -------------------------------------------------
            workbook.Save("RangeDemo.xlsx");
        }
    }
}