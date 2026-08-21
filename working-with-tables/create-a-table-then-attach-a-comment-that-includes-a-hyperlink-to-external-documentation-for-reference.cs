// Title: Add an Excel Table with a Visible Comment and Hyperlink using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, populate a range, convert it into a ListObject (table), insert a visible comment on the header cell, embed an HTML hyperlink to the Aspose.Cells documentation, and save the file as TableWithComment.xlsx with Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# table comment hyperlink | Add visible comment Aspose.Cells .NET | ListObject with HtmlNote comment | Excel comment hyperlink Aspose | Save workbook with table and comment
// Common Searches: Aspose.Cells add comment with link C# | Create Excel table and attach comment Aspose | Visible comment containing hyperlink Aspose.Cells | How to use HtmlNote in Aspose.Cells comment | C# ListObject comment hyperlink example
// Developer Intent: Create a worksheet table and attach a visible comment that includes a clickable link to external documentation.
// Use Cases: Automated reports where column headers link to API reference pages. | Template workbooks that guide end‑users to online usage guidelines via cell comments. | Training materials that embed documentation URLs directly in Excel comments for quick access.
// AI Prompts: Generate C# code to add a ListObject and a visible comment with an HTML hyperlink using Aspose.Cells. | Show how to set the HtmlNote property of a comment so the link opens in a browser when clicked. | Explain steps to embed a documentation URL in a table header comment and save the workbook as XLSX.

using System;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // Demonstrates how to create a workbook, populate a range, convert it into a ListObject (table), insert a visible comment on the header cell, embed an HTML hyperlink to the Aspose.Cells documentation, and save the file as TableWithComment.xlsx with Aspose.Cells for .NET.
    class TableWithCommentDemo
    {
        static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
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
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the table
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");

            // Create a ListObject (table) that covers the data range A1:B3
            int tableIndex = sheet.ListObjects.Add(0, 0, 2, 1, true);
            var listObj = sheet.ListObjects[tableIndex];
            // If needed, set a display name (property may vary by version)
            // listObj.DisplayName = "SampleTable";

            // Add a comment to the header cell A1
            int commentIdx = sheet.Comments.Add("A1");
            var comment = sheet.Comments[commentIdx];

            // Embed a hyperlink to external documentation using HtmlNote
            comment.HtmlNote = "<a href=\"https://docs.aspose.com/cells/net/\" target=\"_blank\">Aspose.Cells Documentation</a>";
            comment.IsVisible = true;

            // Optionally set a plain‑text comment for the table (if supported)
            // listObj.Comment = "See attached comment for documentation link.";

            // Save the workbook
            string outputPath = "TableWithComment.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }
}
