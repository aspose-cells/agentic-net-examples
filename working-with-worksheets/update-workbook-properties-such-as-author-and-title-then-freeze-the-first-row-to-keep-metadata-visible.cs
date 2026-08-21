// Title: Aspose.Cells for .NET – Set Author & Title Metadata and Freeze the Top Row
// Description: Demonstrates how to assign the built‑in Author and Title properties of a new Workbook, apply FreezePanes to keep the first row static, and save the file as UpdatedWorkbook.xlsx using C#.
// Keywords: Aspose.Cells C# | set workbook author | set workbook title | freeze top row | FreezePanes example | document properties .NET | Excel metadata Aspose
// Common Searches: Aspose.Cells set author property C# | How to freeze the first row in Aspose.Cells | Update Excel metadata with Aspose.Cells .NET | FreezePanes usage Aspose.Cells example | Add title to workbook using Aspose.Cells
// Developer Intent: Add author and title metadata to an Excel workbook and lock the header row in place before saving.
// Use Cases: Create a branded report that includes author and title information while keeping column headings visible. | Build a data‑entry template where the top row stays fixed for easier navigation. | Export a spreadsheet with embedded metadata for compliance or auditing purposes.
// AI Prompts: Show C# code that sets Author, Title, and Subject properties and freezes the first two rows with Aspose.Cells. | Explain the parameters of Worksheet.FreezePanes and how to keep header rows static in an Excel file. | Provide an Aspose.Cells example that updates multiple built‑in properties and applies FreezePanes to a specific worksheet.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsExamples
{
    // Demonstrates how to assign the built‑in Author and Title properties of a new Workbook, apply FreezePanes to keep the first row static, and save the file as UpdatedWorkbook.xlsx using C#.
    public class UpdatePropertiesAndFreezeRow
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set built‑in document properties
            DocumentProperty authorProp = workbook.BuiltInDocumentProperties["Author"];
            authorProp.Value = "John Smith";

            DocumentProperty titleProp = workbook.BuiltInDocumentProperties["Title"];
            titleProp.Value = "Sample Workbook";

            // Freeze the first row of the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            // FreezePanes(row, column, totalRows, totalColumns) – freeze 1 row, 0 columns
            sheet.FreezePanes(1, 0, 1, 0);

            // Save the workbook
            workbook.Save("UpdatedWorkbook.xlsx");
        }
    }
}
