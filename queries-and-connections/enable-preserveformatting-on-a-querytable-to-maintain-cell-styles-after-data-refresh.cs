// Title: Aspose.Cells for .NET – PreserveFormatting on QueryTable to retain Excel cell styles after refresh (C#)
// Description: Demonstrates how to create a workbook, style a header row, add a ListObject (Excel table) and enable the QueryTable.PreserveFormatting flag so that formatting survives data refreshes. The example saves the file as QueryTablePreserveFormattingDemo.xlsx.
// Keywords: Aspose.Cells | C# | .NET | QueryTable PreserveFormatting | Excel ListObject formatting | keep cell style after refresh | Excel table refresh Aspose | preserve header formatting | automated data import Excel | Aspose.Cells example
// Common Searches: Aspose.Cells preserve formatting on querytable | C# set QueryTable.PreserveFormatting true | keep Excel header style after refresh Aspose | listobject querytable formatting .NET | how to retain cell styles when refreshing a query table
// Developer Intent: Enable the PreserveFormatting property on a QueryTable so that existing cell styles are not overwritten when the table data is refreshed.
// Use Cases: Generate reports with styled headers that must stay unchanged during periodic data imports. | Automate Excel data refreshes while preserving custom cell formatting in enterprise dashboards. | Create reusable Excel templates where ListObject QueryTables retain user‑defined styles after each update.
// AI Prompts: Show C# code using Aspose.Cells to set QueryTable.PreserveFormatting = true and verify that header formatting remains after a data refresh. | Explain the effect of the PreserveFormatting flag on a QueryTable in Aspose.Cells and how it interacts with ListObject styling. | Provide a step‑by‑step guide to create a ListObject with a QueryTable in Aspose.Cells and keep custom cell styles during refresh operations.

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.Drawing;
using System.IO;

// Demonstrates how to create a workbook, style a header row, add a ListObject (Excel table) and enable the QueryTable.PreserveFormatting flag so that formatting survives data refreshes. The example saves the file as QueryTablePreserveFormattingDemo.xlsx.
class QueryTablePreserveFormattingDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data that will be used as the source for the table
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Price");
            worksheet.Cells["A2"].PutValue("Laptop");
            worksheet.Cells["B2"].PutValue(1200);
            worksheet.Cells["A3"].PutValue("Phone");
            worksheet.Cells["B3"].PutValue(800);

            // Apply a style to the header row (A1:B1)
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.ForegroundColor = Color.LightBlue;
            headerStyle.Pattern = BackgroundType.Solid;
            worksheet.Cells["A1"].SetStyle(headerStyle);
            worksheet.Cells["B1"].SetStyle(headerStyle);

            // Add a ListObject (Excel table) to the range A1:B3
            int tableIndex = worksheet.ListObjects.Add("A1", "B3", true);
            ListObject table = worksheet.ListObjects[tableIndex];
            table.DisplayName = "DemoTable";

            // Enable preserving formatting when the table is refreshed
            if (table.QueryTable != null)
            {
                table.QueryTable.PreserveFormatting = true;
            }

            // Define output file path
            string outputPath = "QueryTablePreserveFormattingDemo.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
