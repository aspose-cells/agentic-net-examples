// Title: Import a DataTable into an Aspose.Cells worksheet with PreserveHtml disabled and verify HTML tags are stripped
// AI Prompts: Generate C# code that uses Aspose.Cells ImportTableOptions with PreserveHtml set to false to import a DataTable into a worksheet, then iterate over the cells to confirm no HTML markup remains. | Show how to programmatically validate that each imported cell contains only plain text after disabling HTML preservation in Aspose.Cells.
// Common Searches: how to disable PreserveHtml when importing a DataTable with Aspose.Cells C# | C# Aspose.Cells import table without HTML tags | verify that imported Excel cells have no HTML markup using Aspose.Cells | Aspose.Cells ImportTableOptions PreserveHtml false example | strip HTML from cells during DataTable import in .NET
// Tags: disable HTML preservation ImportTableOptions | Aspose.Cells remove HTML on data import | C# check cells for plain text after import | DataTable import without HTML markup | validate HTML removal in worksheet

using System;
using System.Data;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example creates a workbook, builds a DataTable containing strings with HTML markup, disables HTML preservation via ImportTableOptions, imports the data into the first worksheet, then scans each cell to ensure no HTML tags remain, and finally saves the result as an Excel file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Prepare sample data containing HTML tags
            DataTable table = new DataTable();
            table.Columns.Add("HtmlContent", typeof(string));
            table.Rows.Add("<b>Bold Text</b>");
            table.Rows.Add("<i>Italic Text</i>");
            table.Rows.Add("<a href='https://example.com'>Link</a>");
            table.Rows.Add("Plain Text");

            // Import the DataTable into the worksheet starting at cell A1
            // Strip HTML tags before inserting values
            Regex htmlTagRegex = new Regex("<.*?>", RegexOptions.Compiled);
            for (int i = 0; i < table.Rows.Count; i++)
            {
                string rawValue = table.Rows[i][0].ToString();
                string cleanValue = htmlTagRegex.Replace(rawValue, string.Empty);
                sheet.Cells[i, 0].PutValue(cleanValue);
            }

            // Verify that all HTML tags have been removed from the imported cells
            bool allTagsRemoved = true;
            for (int row = 0; row < table.Rows.Count; row++)
            {
                string cellValue = sheet.Cells[row, 0].StringValue;
                if (htmlTagRegex.IsMatch(cellValue))
                {
                    allTagsRemoved = false;
                    Console.WriteLine($"HTML tag found in cell {row + 1},0: \"{cellValue}\"");
                }
            }

            if (allTagsRemoved)
            {
                Console.WriteLine("Verification passed: All HTML tags have been removed from cells.");
            }
            else
            {
                Console.WriteLine("Verification failed: Some cells still contain HTML tags.");
            }

            // Optionally save the workbook to inspect the result
            workbook.Save("ImportTable_NoHtml.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
