// Title: C# – Create a self‑expanding named range that grows with added columns using Aspose.Cells
// Description: This example builds a workbook, adds a named range called ExpandingRange that starts at A1 and automatically extends to the last non‑empty cell in row 1 using the INDEX‑COUNTA formula, shows the range address before and after inserting a new column, and saves the file as ExpandingNamedRange.xlsx.
// Keywords: Aspose.Cells | C# | .NET | named range | dynamic range | auto‑expand columns | INDEX function | COUNTA function | Excel automation | workbook manipulation | self‑adjusting range
// Common Searches: Aspose.Cells dynamic named range C# | expand named range when inserting column | auto expanding range Excel using Aspose | C# INDEX COUNTA named range example | how to create a self‑adjusting range in Aspose.Cells
// Developer Intent: Define a named range that automatically extends to the right as new columns are inserted.
// Use Cases: Maintain a header range that grows when additional data columns are added. | Provide a chart data source that updates automatically with new columns. | Apply formulas, formatting, or data validation across a column set that can expand over time.
// AI Prompts: Generate C# Aspose.Cells code to create a named range that expands right when columns are added, using the INDEX‑COUNTA formula. | Show how to retrieve and display the address of an expanding named range before and after inserting a column. | Explain how the formula =Sheet1!$A$1:INDEX(Sheet1!$1:$1,1,COUNTA(Sheet1!$1:$1)) creates a self‑expanding range in Aspose.Cells.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsNamedRangeExpansion
{
    // This example builds a workbook, adds a named range called ExpandingRange that starts at A1 and automatically extends to the last non‑empty cell in row 1 using the INDEX‑COUNTA formula, shows the range address before and after inserting a new column, and saves the file as ExpandingNamedRange.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate initial data in row 1 (A1, B1, C1)
                cells["A1"].PutValue("Header1");
                cells["B1"].PutValue("Header2");
                cells["C1"].PutValue("Header3");
                cells["A2"].PutValue(10);
                cells["B2"].PutValue(20);
                cells["C2"].PutValue(30);

                // Add a named range that expands to the right as new columns are added.
                // The formula uses INDEX together with COUNTA to determine the last used column in row 1.
                // It starts at A1 and ends at the last non‑empty cell in row 1.
                int nameIndex = workbook.Worksheets.Names.Add("ExpandingRange");
                Name expandingRange = workbook.Worksheets.Names[nameIndex];
                expandingRange.RefersTo = "=Sheet1!$A$1:INDEX(Sheet1!$1:$1,1,COUNTA(Sheet1!$1:$1))";

                // Verify the range before adding a new column
                AsposeRange rangeBefore = expandingRange.GetRange();
                Console.WriteLine("Range before adding column: " + rangeBefore.Address); // Expected: A1:C2

                // Insert a new column at the end (after column C)
                cells.InsertColumn(3); // Inserts column D (zero‑based index)

                // Populate data in the new column D
                cells["D1"].PutValue("Header4");
                cells["D2"].PutValue(40);

                // Retrieve the range again; it should now include the new column D
                AsposeRange rangeAfter = expandingRange.GetRange();
                Console.WriteLine("Range after adding column: " + rangeAfter.Address); // Expected: A1:D2

                // Save the workbook (lifecycle rule: save)
                workbook.Save("ExpandingNamedRange.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
