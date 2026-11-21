// Title: Apply a light‑blue solid fill and bold font to a DataTable‑derived range using Aspose.Cells for .NET
// Common Searches: how to set a light blue solid background for a DataTable range with Aspose.Cells in C# | Aspose.Cells C# example for applying bold font to imported DataTable cells | create a style object and apply it to a worksheet range generated from a DataSet using Aspose.Cells | using StyleFlag.All to format an entire range in Aspose.Cells .NET | C# code to style cells with background color and bold text in an Aspose.Cells workbook
// Tags: apply style to worksheet range Aspose.Cells | solid fill background color Aspose.Cells | bold font formatting Aspose.Cells | styleflag all true Aspose.Cells | import datatable to worksheet Aspose.Cells | create style object C# Aspose.Cells

using System;
using System.Data;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook, imports a DataTable into the first worksheet, defines a style with a light‑blue solid fill and bold font, applies the style to the entire imported range using a StyleFlag, and saves the result as StyledData.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // ----- Sample data creation (replace with your actual DataSet) -----
            DataTable dt = new DataTable();
            dt.Columns.Add("Name");
            dt.Columns.Add("Value", typeof(int));
            dt.Rows.Add("Alpha", 100);
            dt.Rows.Add("Beta", 200);
            dt.Rows.Add("Gamma", 300);
            // -------------------------------------------------------------------

            // Manually import the DataTable into the worksheet starting at cell A1 (including column headers)
            int startRow = 0;
            int startColumn = 0;
            bool importColumnNames = true;

            int currentRow = startRow;

            // Write column headers if required
            if (importColumnNames)
            {
                for (int c = 0; c < dt.Columns.Count; c++)
                {
                    cells[currentRow, startColumn + c].PutValue(dt.Columns[c].ColumnName);
                }
                currentRow++;
            }

            // Write data rows
            foreach (DataRow dr in dt.Rows)
            {
                for (int c = 0; c < dt.Columns.Count; c++)
                {
                    cells[currentRow, startColumn + c].PutValue(dr[c]);
                }
                currentRow++;
            }

            // Determine the range that contains the imported data (including header row)
            int totalRows = dt.Rows.Count + (importColumnNames ? 1 : 0);
            int totalCols = dt.Columns.Count;

            // Resolve ambiguity between Aspose.Cells.Range and System.Range
            Aspose.Cells.Range dataRange = cells.CreateRange(startRow, startColumn, totalRows, totalCols);

            // Create a new style object
            Style style = workbook.CreateStyle();

            // Set solid fill to light blue
            style.ForegroundColor = Color.LightBlue;
            style.Pattern = BackgroundType.Solid;

            // Set font to bold
            style.Font.IsBold = true;

            // Apply the style to the entire range
            StyleFlag flag = new StyleFlag { All = true };
            dataRange.ApplyStyle(style, flag);

            // Save the workbook to a file
            string outputPath = "StyledData.xlsx";

            // Ensure the directory exists (handle possible null from GetDirectoryName)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (string.IsNullOrEmpty(outputDir))
            {
                outputDir = Directory.GetCurrentDirectory();
            }

            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
