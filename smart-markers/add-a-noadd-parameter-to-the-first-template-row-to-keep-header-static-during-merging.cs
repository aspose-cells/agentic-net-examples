// Title: Add a (noadd) parameter to the first template row to keep the header static during smart‑marker merging with Aspose.Cells in C#
// Common Searches: how to prevent a header row from being duplicated when using Aspose.Cells smart markers in C# | C# Aspose.Cells add noadd parameter to a smart marker row | keep first row static during WorkbookDesigner smart marker processing | Aspose.Cells example of using noadd to stop row insertion | merge data into an Excel template without adding new rows using Aspose.Cells
// Tags: smart markers noadd parameter | static header row Aspose.Cells | WorkbookDesigner process smart markers C# | add noadd to template row Aspose.Cells | prevent row insertion during smart marker merge | Aspose.Cells template modification C#

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The example loads a template workbook, scans the first row for smart marker strings, appends the (noadd) parameter to keep that row static during merging, sets a DataTable as the data source, processes the markers with WorkbookDesigner, and saves the merged result.
class MergeWithStaticHeader
{
    static void Main()
    {
        try
        {
            // Path to the template workbook that contains smart markers
            const string templatePath = "Template.xlsx";

            // Verify that the template file exists to avoid FileNotFoundException
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Error: Template file '{templatePath}' not found.");
                return;
            }

            // Load the template workbook
            Workbook workbook = new Workbook(templatePath);

            // Get the first worksheet (assumed to contain the template rows)
            Worksheet sheet = workbook.Worksheets[0];

            // Define the index of the first template row (0‑based). Adjust if the header is on a different row.
            int templateRowIndex = 0;

            // Iterate through all cells in the first template row
            for (int col = 0; col < sheet.Cells.MaxColumn; col++)
            {
                Cell cell = sheet.Cells[templateRowIndex, col];
                if (cell.Type == CellValueType.IsString && cell.StringValue.Contains("&="))
                {
                    // Example marker: &=Data.Column
                    // Append "(noadd)" to keep this row static during merging
                    string marker = cell.StringValue;

                    // If the marker already contains parameters, insert noadd before the closing parenthesis
                    if (marker.Contains("(") && marker.EndsWith(")"))
                    {
                        int insertPos = marker.LastIndexOf(')');
                        marker = marker.Insert(insertPos, ",noadd");
                    }
                    else
                    {
                        // No existing parameters – simply add (noadd)
                        marker = marker + "(noadd)";
                    }

                    cell.PutValue(marker);
                }
            }

            // Prepare the data source (example DataTable)
            DataTable dt = new DataTable("Data");
            dt.Columns.Add("Column", typeof(string));
            dt.Rows.Add("Value 1");
            dt.Rows.Add("Value 2");
            dt.Rows.Add("Value 3");

            // Process smart markers – the first row will stay static because of the noadd parameter
            try
            {
                // Use WorkbookDesigner to process smart markers
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource(dt);
                designer.Process();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Smart marker processing failed: {ex.Message}");
                return;
            }

            // Save the merged workbook
            const string resultPath = "MergedResult.xlsx";
            workbook.Save(resultPath);
            Console.WriteLine($"Merged workbook saved to '{resultPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
