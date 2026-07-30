// Title: Create a named range from a ListObject (Excel table) and use it in formulas – Aspose.Cells for .NET
// Description: This example builds a workbook, adds a ListObject covering A1:B6, extracts the table's DataRange, defines a workbook‑level named range that points to that address, inserts a SUM formula that references the named range, calculates the result, and saves the file.
// Keywords: Aspose.Cells ListObject named range | C# retrieve table address range | Excel table DataRange Aspose | programmatic named range .NET | SUM formula using table range
// Common Searches: Aspose.Cells get ListObject address | create named range from Excel table C# | use table DataRange in formula Aspose | define workbook named range programmatically | sum table range with named range Aspose.Cells
// Developer Intent: Programmatically obtain the address of a ListObject, create a named range that references that address, and use the named range in worksheet formulas.
// Use Cases: Reference a table in multiple calculations without hard‑coding cell coordinates. | Provide a dynamic data source for charts or pivot tables that expands with the table. | Standardize formulas (SUM, AVERAGE, etc.) across worksheets by using a single named range.
// AI Prompts: Generate C# code with Aspose.Cells that adds a ListObject, extracts its DataRange, creates a named range, and applies a SUM formula referencing it. | Explain how to keep a named range synchronized with a ListObject when rows are inserted or deleted using Aspose.Cells. | Show how to use a table‑based named range as the source series for a chart created with Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsTableNamedRangeDemo
{
    // This example builds a workbook, adds a ListObject covering A1:B6, extracts the table's DataRange, defines a workbook‑level named range that points to that address, inserts a SUM formula that references the named range, calculates the result, and saves the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "DataSheet";

                // Populate sample data that will become a table
                sheet.Cells["A1"].PutValue("ID");
                sheet.Cells["B1"].PutValue("Value");
                for (int i = 0; i < 5; i++)
                {
                    sheet.Cells[i + 1, 0].PutValue(i + 1);               // Column A: IDs 1..5
                    sheet.Cells[i + 1, 1].PutValue((i + 1) * 10);      // Column B: Values 10,20,...
                }

                // Add a ListObject (Excel Table) covering the data range A1:B6
                int tableIndex = sheet.ListObjects.Add(0, 0, 5, 1, true);
                ListObject table = sheet.ListObjects[tableIndex];
                table.DisplayName = "MyTable";

                // Retrieve the data range of the table (including header)
                AsposeRange dataRange = table.DataRange;

                // Create a named range that refers to the table's data range
                const string namedRangeName = "TableRange";
                int nameIdx = workbook.Worksheets.Names.Add(namedRangeName);
                Name namedRange = workbook.Worksheets.Names[nameIdx];
                namedRange.RefersTo = $"={sheet.Name}!{dataRange.Address}";

                // Use the named range in a formula (sum of the table range)
                sheet.Cells["D1"].Formula = $"=SUM({namedRangeName})";

                // Calculate formulas so the result is available
                workbook.CalculateFormula();

                // Output the result to console
                Console.WriteLine($"Named range '{namedRangeName}' refers to: {namedRange.RefersTo}");
                Console.WriteLine($"Sum of the table range: {sheet.Cells["D1"].Value}");

                // Save the workbook
                string outputPath = "TableNamedRangeDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
