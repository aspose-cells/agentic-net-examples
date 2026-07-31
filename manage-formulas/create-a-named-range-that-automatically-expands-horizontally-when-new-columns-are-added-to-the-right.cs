// Title: Aspose.Cells for .NET: Create a Horizontally Expanding Named Range with INDEX‑COUNTA (C#)
// Description: This C# example demonstrates how to define a named range that automatically expands across a row as new columns are inserted. It uses the formula =Sheet1!$A$1:INDEX(Sheet1!$1:$1, COUNTA(Sheet1!$1:$1)) to cover all non‑empty cells in the first row, inserts a column, adds data, recalculates formulas, and retrieves the updated range address, showcasing a fully dynamic named range in Aspose.Cells.
// Keywords: Aspose.Cells | C# dynamic named range | horizontal expanding named range | INDEX function | COUNTA function | InsertColumn | Workbook.CalculateFormula | named range address | Excel dynamic range | Aspose.Cells example | GitHub Aspose.Cells | Excel chart data source
// Common Searches: Aspose.Cells create dynamic named range C# | expand named range when inserting columns Aspose.Cells | INDEX COUNTA named range Aspose.Cells .NET | update named range after InsertColumn Aspose.Cells | retrieve named range address after recalculation Aspose.Cells | C# Aspose.Cells example dynamic range GitHub
// Developer Intent: Define a named range that automatically grows to include any new columns added to the right of the worksheet.
// Use Cases: Maintain a header range for monthly data that expands as new month columns are added. | Provide a chart data source that updates automatically when additional columns are inserted. | Apply conditional formatting or data validation to a range that adjusts without manual range edits.
// AI Prompts: Generate C# code using Aspose.Cells that creates a named range expanding horizontally based on non‑empty cells in the first row. | Show how to recalculate formulas and obtain the updated address of a dynamic named range after inserting a column. | Explain why the INDEX‑COUNTA formula enables a horizontally dynamic named range in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsDynamicNamedRangeDemo
{
    // This C# example demonstrates how to define a named range that automatically expands across a row as new columns are inserted. It uses the formula =Sheet1!$A$1:INDEX(Sheet1!$1:$1, COUNTA(Sheet1!$1:$1)) to cover all non‑empty cells in the first row, inserts a column, adds data, recalculates formulas, and retrieves the updated range address, showcasing a fully dynamic named range in Aspose.Cells.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate initial data in row 1 (A1:C1)
                cells["A1"].PutValue("Jan");
                cells["B1"].PutValue("Feb");
                cells["C1"].PutValue("Mar");

                // Create a named range that expands horizontally based on the number of non‑empty cells in row 1
                // Formula: =Sheet1!$A$1:INDEX(Sheet1!$1:$1, COUNTA(Sheet1!$1:$1))
                int nameIndex = workbook.Worksheets.Names.Add("MyRange");
                Name myRangeName = workbook.Worksheets.Names[nameIndex];
                myRangeName.RefersTo = "=Sheet1!$A$1:INDEX(Sheet1!$1:$1, COUNTA(Sheet1!$1:$1))";

                // Retrieve and display the initial range address
                Aspose.Cells.Range initialRange = myRangeName.GetRange();
                Console.WriteLine("Initial named range address: " + initialRange.Address);

                // Insert a new column after column C (which is index 2)
                // This shifts existing columns to the right, making room for a new column D
                cells.InsertColumn(3, true); // Insert at column index 3 (D)

                // Add data to the newly inserted column D (which is now column index 3)
                cells["D1"].PutValue("Apr");

                // Recalculate formulas so that the named range reflects the new column count
                workbook.CalculateFormula();

                // Retrieve and display the updated range address
                Aspose.Cells.Range updatedRange = myRangeName.GetRange();
                Console.WriteLine("Updated named range address after inserting column: " + updatedRange.Address);

                // Save the workbook (optional, just to complete the lifecycle)
                string outputPath = "DynamicNamedRangeDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
