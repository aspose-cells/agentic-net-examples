// Title: C# – Apply Dark2 (Accent2) Theme Color as Default Border for All Aspose.Cells Tables
// Description: This example creates a workbook, adds a sample ListObject, then loops through every worksheet and each table to build a thin border style. The style uses the workbook’s Accent2 (Dark2) theme color for all four borders and is applied only to the table range before saving the file.
// Keywords: Aspose.Cells C# | theme color border | Accent2 Dark2 | table border style | ListObject formatting | StyleFlag ApplyStyle | default table border | thin border Excel | sample code GitHub | Excel workbook styling
// Common Searches: Aspose.Cells set Accent2 theme color for table borders C# | apply thin dark2 borders to all ListObjects in a workbook | default border style for tables using Aspose.Cells .NET | how to use ThemeColor with StyleFlag in Aspose.Cells
// Developer Intent: Set the Dark2 (Accent2) theme color as the default border for every table in a workbook.
// Use Cases: Generate corporate reports where all tables share a consistent Dark2 border for branding. | Automate styling of existing worksheets so every ListObject inherits the workbook’s Accent2 border. | Create a reusable template that enforces thin, theme‑colored borders on new tables before export.
// AI Prompts: Provide C# code that iterates over all ListObjects in an Aspose.Cells workbook and applies a thin border using the Accent2 (Dark2) theme color. | Show how to use StyleFlag to apply only border formatting to a table range while preserving other cell styles. | Explain how to set ThemeColor for table borders so that new tables automatically use the workbook’s default Accent2 color.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables; // For ListObject
using AsposeRange = Aspose.Cells.Range; // Resolve ambiguity with System.Range

namespace AsposeCellsThemeBorderForTables
{
    // This example creates a workbook, adds a sample ListObject, then loops through every worksheet and each table to build a thin border style. The style uses the workbook’s Accent2 (Dark2) theme color for all four borders and is applied only to the table range before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet (add more if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the table
                worksheet.Cells["A1"].PutValue("Header1");
                worksheet.Cells["B1"].PutValue("Header2");
                worksheet.Cells["A2"].PutValue("Data1");
                worksheet.Cells["B2"].PutValue("Data2");
                worksheet.Cells["A3"].PutValue("Data3");
                worksheet.Cells["B3"].PutValue("Data4");

                // Add a table covering the data range (including header)
                int tableIndex = worksheet.ListObjects.Add(0, 0, 2, 1, true);
                ListObject table = worksheet.ListObjects[tableIndex];
                table.DisplayName = "SampleTable";

                // Iterate through all worksheets
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    // Iterate through all tables (ListObjects) in the worksheet
                    foreach (ListObject lo in ws.ListObjects)
                    {
                        // Determine the full range of the table (including header)
                        int startRow = lo.StartRow;
                        int startColumn = lo.StartColumn;
                        int totalRows = lo.EndRow - lo.StartRow + 1;
                        int totalColumns = lo.EndColumn - lo.StartColumn + 1;
                        AsposeRange tableRange = ws.Cells.CreateRange(startRow, startColumn, totalRows, totalColumns);

                        // Create a new style for the table borders
                        Style borderStyle = workbook.CreateStyle();

                        // Set thin line style for all borders
                        borderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                        borderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                        borderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                        borderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

                        // Use the workbook's Accent2 (Dark2) theme color
                        ThemeColor dark2Theme = new ThemeColor(ThemeColorType.Accent2, 0);
                        borderStyle.Borders[BorderType.TopBorder].ThemeColor = dark2Theme;
                        borderStyle.Borders[BorderType.BottomBorder].ThemeColor = dark2Theme;
                        borderStyle.Borders[BorderType.LeftBorder].ThemeColor = dark2Theme;
                        borderStyle.Borders[BorderType.RightBorder].ThemeColor = dark2Theme;

                        // Apply only border settings
                        StyleFlag flag = new StyleFlag { Borders = true };
                        tableRange.ApplyStyle(borderStyle, flag);
                    }
                }

                // Save the workbook
                workbook.Save("TablesWithDark2Border.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
