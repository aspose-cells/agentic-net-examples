// Title: Aspose.Cells for .NET – Apply Dark2 Theme Accent as Default Thin Border to All ListObject Tables
// Description: C# example that creates a workbook, adds a ListObject table, then loops through every worksheet and each table to apply a thin border style whose ThemeColor is set to the workbook's Accent2 (Dark2) color. Only border formatting is changed, and the workbook is saved with the new styling.
// Keywords: Aspose.Cells C# table border theme | Dark2 theme accent border | Accent2 thin border Aspose.Cells | apply theme color to ListObject borders | set default table border Aspose.Cells | C# Excel theme color borders | Aspose.Cells style flag borders
// Common Searches: Aspose.Cells set table border to Dark2 theme color | C# apply theme accent to Excel table borders | How to use ThemeColor Accent2 for ListObject borders in Aspose.Cells | Apply thin borders to all tables in a workbook using Aspose.Cells | Default border color for tables Aspose.Cells .NET
// Developer Intent: Use Aspose.Cells to make the workbook's Dark2 (Accent2) theme color the default thin border for every ListObject table.
// Use Cases: Generate reports where all tables share a consistent Dark2 border that matches the workbook theme. | Standardize the appearance of tables across multiple worksheets in automated Excel exports. | Retrofit existing workbooks by programmatically updating each table’s border to the theme’s Accent2 color.
// AI Prompts: Write C# code with Aspose.Cells that sets the Dark2 (Accent2) theme color as a thin border for all ListObject tables in a workbook. | Show how to change the border color to a different theme accent (e.g., Accent3) while keeping the same thin style. | Explain how to apply the Dark2 border style only to data rows of each table, leaving header borders unchanged.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsThemeBorderDemo
{
    // C# example that creates a workbook, adds a ListObject table, then loops through every worksheet and each table to apply a thin border style whose ThemeColor is set to the workbook's Accent2 (Dark2) color. Only border formatting is changed, and the workbook is saved with the new styling.
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

                // Populate some data and create a table (ListObject)
                sheet.Cells["A1"].PutValue("Header1");
                sheet.Cells["B1"].PutValue("Header2");
                sheet.Cells["A2"].PutValue(10);
                sheet.Cells["B2"].PutValue(20);

                // Add a table that includes the header row
                int tableIndex = sheet.ListObjects.Add(0, 0, 2, 2, true);
                ListObject table = sheet.ListObjects[tableIndex];

                // Iterate through all worksheets and their tables
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    foreach (ListObject lo in ws.ListObjects)
                    {
                        // Build a range that covers the whole table (including header)
                        int firstRow = lo.StartRow;
                        int firstCol = lo.StartColumn;
                        int rowCount = lo.EndRow - lo.StartRow + 1;
                        int colCount = lo.EndColumn - lo.StartColumn + 1;
                        AsposeRange tableRange = ws.Cells.CreateRange(firstRow, firstCol, rowCount, colCount);

                        // Create a new style
                        Style style = workbook.CreateStyle();

                        // Use the theme's Accent2 (Dark2) color for thin borders
                        ThemeColor dark2Theme = new ThemeColor(ThemeColorType.Accent2, 0);
                        style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                        style.Borders[BorderType.TopBorder].ThemeColor = dark2Theme;

                        style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                        style.Borders[BorderType.BottomBorder].ThemeColor = dark2Theme;

                        style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                        style.Borders[BorderType.LeftBorder].ThemeColor = dark2Theme;

                        style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
                        style.Borders[BorderType.RightBorder].ThemeColor = dark2Theme;

                        // Apply only border formatting
                        StyleFlag flag = new StyleFlag { Borders = true };

                        // Apply the style to the table range
                        tableRange.ApplyStyle(style, flag);
                    }
                }

                // Define output file path
                string outputPath = "TablesWithDark2Border.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
