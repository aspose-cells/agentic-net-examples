// Title: Apply a built‑in table style to an Aspose.Cells ListObject while preserving custom header formatting in C#
// Common Searches: Aspose.Cells C# apply TableStyleLight1 to a ListObject without overwriting header style | preserve custom header formatting when setting TableStyleType in Aspose.Cells | create Excel table with predefined style using Aspose.Cells .NET | how to keep header cell colors after applying a built‑in table style in Aspose.Cells | C# add ListObject to worksheet and retain header formatting with Aspose.Cells
// Tags: apply built‑in table style Aspose.Cells | C# ListObject custom header formatting | TableStyleType TableStyleLight1 usage | Excel table styling with Aspose.Cells .NET | preserve cell style after applying table style

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a workbook, fills it with sample data, applies a bold light‑gray style to the header row, adds a ListObject covering the data range, sets the predefined TableStyleLight1 on the table, and preserves the custom header formatting before saving the file as TableWithStyle.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["C2"].PutValue(85);
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");
            sheet.Cells["C3"].PutValue(92);
            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue("Charlie");
            sheet.Cells["C4"].PutValue(78);

            // Apply custom formatting to header row (preserve this formatting)
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.ForegroundColor = Color.LightGray;
            headerStyle.Pattern = BackgroundType.Solid;
            StyleFlag headerFlag = new StyleFlag { All = true };
            sheet.Cells.CreateRange("A1:C1").ApplyStyle(headerStyle, headerFlag);

            // Define the range that will become a table (zero‑based indices)
            CellArea tableArea = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 3,
                EndColumn = 2
            };

            // Add a ListObject (table) to the worksheet
            int tableIndex = sheet.ListObjects.Add(
                tableArea.StartRow,
                tableArea.StartColumn,
                tableArea.EndRow,
                tableArea.EndColumn,
                true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Set a display name for the table
            table.DisplayName = "SampleTable";

            // Apply a predefined table style (choose an existing style)
            table.TableStyleType = TableStyleType.TableStyleLight1;

            // Determine output path and ensure the directory exists
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWithStyle.xlsx");
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
