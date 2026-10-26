// Title: Refresh Excel worksheet data while keeping number formats and cell colors with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# preserve cell formatting when updating worksheet data | keep number format and background color after importing DataTable into Excel with Aspose.Cells | refresh Excel data programmatically without losing custom cell styles using Aspose.Cells .NET
// Tags: capture and restore cell style Aspose.Cells | import DataTable without losing formatting Aspose.Cells | retain numeric display format during worksheet refresh Aspose.Cells | preserve cell fill color when updating Excel Aspose.Cells | reapply original cell styles after data import Aspose.Cells

using Aspose.Cells;
using System;
using System.Data;
using System.IO;

// The example loads an existing workbook (or creates one), captures each cell's current Style, imports refreshed DataTable values into the same range without altering formatting, then reapplies the saved styles to retain number formats and cell colors before saving the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Load existing workbook if it exists; otherwise create a new one.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                // Add a default worksheet.
                workbook.Worksheets.Add("Sheet1");
            }

            // Obtain refreshed data from an external source.
            DataTable refreshedData = GetRefreshedData();

            // Target worksheet (first sheet in this example).
            Worksheet sheet = workbook.Worksheets[0];

            // Define where the data starts (row 0, column 0 in this example).
            int startRow = 0;
            int startColumn = 0;

            int rows = refreshedData.Rows.Count;
            int cols = refreshedData.Columns.Count;

            // ------------------------------------------------------------
            // Preserve formatting: capture the current style of each cell.
            // ------------------------------------------------------------
            Style[,] originalStyles = new Style[rows, cols];
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    Cell cell = sheet.Cells[startRow + i, startColumn + j];
                    originalStyles[i, j] = cell.GetStyle();
                }
            }

            // ------------------------------------------------------------
            // Refresh data: import new values without affecting formatting.
            // ------------------------------------------------------------
            // Aspose.Cells may not have ImportDataTable in some versions, so we fill manually.
            for (int i = 0; i < rows; i++)
            {
                DataRow dr = refreshedData.Rows[i];
                for (int j = 0; j < cols; j++)
                {
                    Cell cell = sheet.Cells[startRow + i, startColumn + j];
                    cell.PutValue(dr[j]);
                }
            }

            // ------------------------------------------------------------
            // Reapply the captured styles to keep number formats and colors.
            // ------------------------------------------------------------
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    Cell cell = sheet.Cells[startRow + i, startColumn + j];
                    cell.SetStyle(originalStyles[i, j]);
                }
            }

            // Save the workbook with refreshed data while preserving formatting.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Mock method that simulates obtaining refreshed data.
    static DataTable GetRefreshedData()
    {
        DataTable dt = new DataTable();
        dt.Columns.Add("Item", typeof(string));
        dt.Columns.Add("Quantity", typeof(int));
        dt.Columns.Add("Price", typeof(double));

        dt.Rows.Add("Apple", 150, 0.55);
        dt.Rows.Add("Banana", 95, 0.30);
        dt.Rows.Add("Cherry", 70, 1.20);

        return dt;
    }
}
