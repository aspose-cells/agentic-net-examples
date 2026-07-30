// Title: Add a comment and hyperlink to an Excel table header using Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, define a ListObject (table) over A1:B3, set its display name, place a comment on header cell A1, attach a hyperlink to the Aspose.Cells online documentation, add a plain‑text comment to the table itself, and save the result as an XLSX file.
// Keywords: Aspose.Cells C# table comment | Aspose.Cells add hyperlink | ListObject comment Aspose | Excel table header comment .NET | Aspose.Cells external documentation link | C# Aspose.Cells create table | Aspose.Cells save as xlsx
// Common Searches: How to add a comment to a table header in Aspose.Cells C# | Aspose.Cells add hyperlink to cell | Create ListObject with comment Aspose.Cells | C# Aspose.Cells embed documentation link | Add comment and hyperlink to Excel using Aspose.Cells
// Developer Intent: Programmatically add a descriptive comment and a clickable link to the header of an Excel table so users can quickly access online documentation.
// Use Cases: Generate reports where each table header contains a comment linking to the API reference for instant user guidance. | Provide end‑users with direct access to the latest Aspose.Cells documentation by embedding hyperlinks in key worksheet cells. | Add audit‑ready comments to tables while linking to external standards or policy documents for compliance tracking.
// AI Prompts: Give me C# code that adds a comment and a hyperlink to a specific cell in an Aspose.Cells workbook. | Show an example that creates a ListObject, sets its display name, adds a header comment, and inserts a hyperlink to the Aspose.Cells documentation. | Explain how to retrieve, update, or delete a comment or hyperlink from a table header after the workbook has been saved.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, define a ListObject (table) over A1:B3, set its display name, place a comment on header cell A1, attach a hyperlink to the Aspose.Cells online documentation, add a plain‑text comment to the table itself, and save the result as an XLSX file.
    public class TableWithCommentAndHyperlink
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
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the table (including header row)
            worksheet.Cells["A1"].PutValue("ID");
            worksheet.Cells["B1"].PutValue("Name");
            worksheet.Cells["A2"].PutValue(1);
            worksheet.Cells["B2"].PutValue("Alice");
            worksheet.Cells["A3"].PutValue(2);
            worksheet.Cells["B3"].PutValue("Bob");

            // Create a ListObject (table) that covers the data range A1:B3
            // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
            int tableIndex = worksheet.ListObjects.Add(0, 0, 3, 2, true);
            ListObject table = worksheet.ListObjects[tableIndex];

            // Set a display name for the table (Name property not available in this version)
            table.DisplayName = "SampleTable";

            // Add a comment to the header cell A1 that references external documentation
            int commentIdx = worksheet.Comments.Add(0, 0); // Row 0, Column 0 => A1
            Comment comment = worksheet.Comments[commentIdx];
            comment.Note = "For more details see the Aspose.Cells documentation.";
            comment.Author = "Developer";

            // Add a hyperlink to the same cell (A1) pointing to the external documentation
            // Parameters: firstRow, firstColumn, totalRows, totalColumns, address
            worksheet.Hyperlinks.Add(0, 0, 1, 1, "https://docs.aspose.com/cells/net/");

            // Optionally, set the ListObject's own comment property (plain text)
            table.Comment = "Table created. See cell A1 for documentation link.";

            // Save the workbook
            string outputPath = "TableWithCommentAndHyperlink.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
