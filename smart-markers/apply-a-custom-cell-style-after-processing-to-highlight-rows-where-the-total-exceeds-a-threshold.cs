// Title: Highlight Excel rows with totals above a threshold using Aspose.Cells for .NET (C#)
// Common Searches: c# aspose.cells highlight rows where column value exceeds 1000 | apply custom yellow style to entire row based on numeric threshold in Aspose.Cells | how to programmatically set background color for rows with total greater than a limit using Aspose.Cells .NET | conditional formatting rows by value with Aspose.Cells C# example | load an Excel workbook with optional password and process rows in Aspose.Cells
// Tags: row highlighting based on numeric threshold Aspose.Cells | apply custom style to entire row Aspose.Cells .NET | conditional formatting using Aspose.Cells C# | set background color for Excel rows Aspose.Cells | load workbook with optional password Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel file (or creates one with sample data), defines a numeric threshold, creates a yellow solid‑fill style, iterates through each data row checking the total in column E, and applies the style to every cell in rows where the total exceeds the threshold before saving the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        Workbook workbook = null;

        try
        {
            // Load existing workbook if it exists; otherwise create a new one with sample data
            if (File.Exists(inputPath))
            {
                try
                {
                    // Attempt to load without password
                    workbook = new Workbook(inputPath);
                }
                catch (CellsException ex) when (ex.Message.Contains("password"))
                {
                    // If the file is password‑protected, try loading with an empty password
                    var loadOptions = new LoadOptions(LoadFormat.Xlsx) { Password = string.Empty };
                    workbook = new Workbook(inputPath, loadOptions);
                }
            }
            else
            {
                // Create a new workbook with sample data
                workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells sheetCells = sheet.Cells;

                // Header
                sheetCells[0, 0].PutValue("Item");
                sheetCells[0, 4].PutValue("Total");

                // Sample rows
                sheetCells[1, 0].PutValue("A");
                sheetCells[1, 4].PutValue(500);
                sheetCells[2, 0].PutValue("B");
                sheetCells[2, 4].PutValue(1500);
            }

            if (workbook == null)
                throw new InvalidOperationException("Workbook could not be initialized.");

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Threshold for highlighting
            double threshold = 1000.0;

            // Style for rows that exceed the threshold
            Style highlightStyle = workbook.CreateStyle();
            highlightStyle.ForegroundColor = Color.Yellow;
            highlightStyle.Pattern = BackgroundType.Solid;

            // Column E (zero‑based index 4) holds totals
            int totalColumnIndex = 4;
            int startRow = 1; // skip header
            int maxRow = cells.MaxDataRow;

            // Apply highlighting where total > threshold
            for (int row = startRow; row <= maxRow; row++)
            {
                Cell totalCell = cells[row, totalColumnIndex];
                double totalValue = 0;

                if (totalCell.Type == CellValueType.IsNumeric)
                {
                    totalValue = totalCell.DoubleValue;
                }
                else
                {
                    double.TryParse(totalCell.StringValue, out totalValue);
                }

                if (totalValue > threshold)
                {
                    for (int col = 0; col <= cells.MaxDataColumn; col++)
                    {
                        cells[row, col].SetStyle(highlightStyle);
                    }
                }
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook processed and saved to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
