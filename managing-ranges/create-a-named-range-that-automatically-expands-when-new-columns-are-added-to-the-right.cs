// Title: Aspose.Cells for .NET – Create a dynamically expanding named range that grows with added columns
// Description: Demonstrates how to define a named range in a workbook using an OFFSET‑COUNTA formula, retrieve its dimensions before and after inserting a column, force formula recalculation, and save the file. The range automatically expands to the right whenever new columns are added, enabling self‑adjusting charts, pivot tables, and formulas.
// Keywords: Aspose.Cells dynamic named range | auto expanding range C# | OFFSET COUNTA formula | named range update after column insert | GetRange true Aspose.Cells | .NET spreadsheet automation | Excel dynamic range programmatically
// Common Searches: Aspose.Cells create named range that expands horizontally | C# OFFSET COUNTA dynamic range example | How to refresh named range after inserting columns Aspose.Cells | GetRange true recalculate formula Aspose.Cells | auto‑adjusting named range for charts in .NET
// Developer Intent: Define a named range that automatically expands when additional columns are inserted into the worksheet.
// Use Cases: Generate a header range that always includes every populated column for chart data sources. | Maintain a self‑updating range for pivot tables or formulas that grow as new data columns are added. | Validate range size programmatically before and after column insertion to ensure downstream calculations stay accurate.
// AI Prompts: Write C# code with Aspose.Cells that creates a named range using OFFSET and COUNTA to auto‑expand horizontally as new columns are added. | Show how to retrieve the updated range after inserting a column and forcing formula recalculation with GetRange(true). | Explain the steps to make a named range self‑adjusting for charts or pivot tables in an Aspose.Cells workbook.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Demonstrates how to define a named range in a workbook using an OFFSET‑COUNTA formula, retrieve its dimensions before and after inserting a column, force formula recalculation, and save the file. The range automatically expands to the right whenever new columns are added, enabling self‑adjusting charts, pivot tables, and formulas.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate header row and some sample data
            worksheet.Cells["A1"].PutValue("Header1");
            worksheet.Cells["B1"].PutValue("Header2");
            worksheet.Cells["C1"].PutValue("Header3");
            worksheet.Cells["A2"].PutValue(1);
            worksheet.Cells["B2"].PutValue(2);
            worksheet.Cells["C2"].PutValue(3);

            // Add a named range that expands automatically to the right
            // The formula uses OFFSET with COUNTA to count filled columns in row 1
            int nameIndex = workbook.Worksheets.Names.Add("MyRange");
            Name myRangeName = workbook.Worksheets.Names[nameIndex];
            myRangeName.RefersTo = "=OFFSET(Sheet1!$A$1,0,0,1,COUNTA(Sheet1!$1:$1))";

            // Retrieve the range before inserting a new column
            AsposeRange rangeBefore = myRangeName.GetRange();
            Console.WriteLine($"Before inserting column: {rangeBefore.Address}, Columns = {rangeBefore.ColumnCount}");

            // Insert a new column to the right of the existing data (after column C)
            worksheet.Cells.InsertColumn(3);
            // Add header and data in the newly inserted column D
            worksheet.Cells["D1"].PutValue("Header4");
            worksheet.Cells["D2"].PutValue(4);

            // Recalculate formulas so the named range reflects the new column count
            workbook.CalculateFormula();

            // Retrieve the range after insertion; use GetRange(true) to force recalculation
            AsposeRange rangeAfter = myRangeName.GetRange(true);
            Console.WriteLine($"After inserting column: {rangeAfter.Address}, Columns = {rangeAfter.ColumnCount}");

            // Save the workbook (ensure the directory exists)
            string outputPath = "DynamicNamedRange.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
