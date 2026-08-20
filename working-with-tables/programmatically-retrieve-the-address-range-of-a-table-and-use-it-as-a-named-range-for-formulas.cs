// Title: Create a Named Range from an Excel Table’s Address Using Aspose.Cells for .NET (C#)
// Description: This example shows how to add a ListObject (Excel table) to a worksheet, retrieve its address range, define a workbook‑level named range that points to that table, and use the named range in a SUM formula with Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# table address | Excel ListObject named range | create named range from table Aspose.Cells | DataRange to named range .NET | Aspose.Cells formula using table column
// Common Searches: Aspose.Cells get ListObject address | define named range from Excel table C# | use table column in SUM formula Aspose.Cells | programmatically create named range in .NET | Aspose.Cells table DataRange example
// Developer Intent: Retrieve the address of an Excel table (ListObject) and create a named range that can be referenced in formulas.
// Use Cases: Simplify formulas by referencing the whole table through a named range. | Perform column‑wise calculations (e.g., SUM, AVERAGE) using the named range. | Ensure formulas stay accurate when the table size changes, as the named range follows the table’s DataRange.
// AI Prompts: Provide C# code that extracts a ListObject's address and creates a named range with Aspose.Cells. | Show an example of adding an Excel table, converting its DataRange to a named range, and using it in a SUM formula. | Explain how to keep a named range synchronized with a table when rows are added or removed in Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables; // For ListObject

namespace AsposeCellsTableToNamedRange
{
    // This example shows how to add a ListObject (Excel table) to a worksheet, retrieve its address range, define a workbook‑level named range that points to that table, and use the named range in a SUM formula with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "DataSheet";

                // Populate sample data for the table (3 columns, 4 rows including header)
                sheet.Cells["A1"].PutValue("ID");
                sheet.Cells["B1"].PutValue("Product");
                sheet.Cells["C1"].PutValue("Price");

                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B2"].PutValue("Apple");
                sheet.Cells["C2"].PutValue(0.5);

                sheet.Cells["A3"].PutValue(2);
                sheet.Cells["B3"].PutValue("Banana");
                sheet.Cells["C3"].PutValue(0.3);

                sheet.Cells["A4"].PutValue(3);
                sheet.Cells["B4"].PutValue("Cherry");
                sheet.Cells["C4"].PutValue(0.8);

                // Add a ListObject (Excel Table) covering the data range A1:C4
                int tableIndex = sheet.ListObjects.Add(0, 0, 3, 2, true);
                ListObject table = sheet.ListObjects[tableIndex];
                table.DisplayName = "ProductTable";

                // Retrieve the data range of the table (including header)
                Aspose.Cells.Range dataRange = table.DataRange;

                // Create a named range that refers to the table's data range
                int nameIndex = workbook.Worksheets.Names.Add("ProductData");
                Name namedRange = workbook.Worksheets.Names[nameIndex];
                // RefersTo must start with '=' and include sheet name
                namedRange.RefersTo = $"={sheet.Name}!{dataRange.Address}";

                // Use the named range in a formula (e.g., total price)
                sheet.Cells["E1"].Formula = "=SUM(ProductData[Price])";

                // Calculate formulas to get the result
                workbook.CalculateFormula();

                // Output the result to console (optional)
                Console.WriteLine($"Sum of ProductData range ({dataRange.Address}) = {sheet.Cells["E1"].Value}");

                // Save the workbook
                string outputPath = "TableToNamedRange.xlsx";
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
