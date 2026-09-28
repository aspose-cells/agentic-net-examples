// Title: Set Author and Title built‑in document properties and freeze the first worksheet row with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that assigns "John Smith" to the Author property and "Sample Workbook" to the Title property of a Workbook, then freezes the top row of the first worksheet using Aspose.Cells. | Write a C# example that updates built‑in document properties of an Excel file and applies FreezePanes to keep the header row visible while scrolling.
// Common Searches: Aspose.Cells C# set workbook author and title metadata | freeze header row in Excel using Aspose.Cells FreezePanes C# | how to update built‑in document properties in an Aspose.Cells workbook | C# example to freeze first row after modifying workbook properties with Aspose.Cells
// Tags: update built‑in document properties Aspose.Cells C# | freeze top row worksheet Aspose.Cells | set author metadata Excel .NET | apply FreezePanes method Aspose.Cells | initialize workbook and modify properties Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsDemo
{
    // Creates a new Workbook, sets the Author and Title built‑in document properties, freezes the first row of the first worksheet, and saves the file as UpdatedWorkbook.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle create)
                Workbook workbook = new Workbook();

                // Update built‑in document properties (author and title)
                DocumentProperty authorProp = workbook.BuiltInDocumentProperties["Author"];
                authorProp.Value = "John Smith";

                DocumentProperty titleProp = workbook.BuiltInDocumentProperties["Title"];
                titleProp.Value = "Sample Workbook";

                // Freeze the first row of the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                // FreezePanes(row, column, totalRows, totalColumns)
                // Freeze rows above 'row' index and columns left of 'column' index
                sheet.FreezePanes(1, 0, 0, 0); // freezes the top row

                // Save the workbook (lifecycle save)
                workbook.Save("UpdatedWorkbook.xlsx");

                // Optional: display confirmation
                Console.WriteLine("Workbook saved with updated properties and first row frozen.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
