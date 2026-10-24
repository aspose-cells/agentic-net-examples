// Title: Apply a strikethrough font style to rows marked as Completed in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# apply strikethrough to rows where status equals Completed | How to format completed tasks with a strikeout font in an Excel file using Aspose.Cells .NET | Programmatically set strikeout font for an entire row in Excel with Aspose.Cells API | C# loop through worksheet rows and apply a custom style based on cell value using Aspose.Cells | Create and reuse a strikeout style for Excel rows in a .NET application with Aspose.Cells
// Tags: strikethrough row formatting Aspose.Cells | custom style IsStrikeout Excel .NET | conditional row styling based on cell value Aspose.Cells | apply style to entire Excel row C# | Excel strikeout font using Aspose.Cells API | reusable workbook style Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads or creates an Excel workbook, defines a custom style with the IsStrikeout font property, scans each data row for the value "Completed" in the Status column, and applies the strikeout style to every cell in those rows before saving the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputFile = "Tasks.xlsx";
            const string outputFile = "Tasks_Updated.xlsx";

            // Ensure the input workbook exists; create a simple one if missing
            if (!File.Exists(inputFile))
            {
                var sampleWb = new Workbook();
                var ws = sampleWb.Worksheets[0];
                var sampleCells = ws.Cells; // renamed to avoid conflict
                sampleCells[0, 0].PutValue("Task");
                sampleCells[0, 1].PutValue("Status");
                sampleCells[1, 0].PutValue("Task 1");
                sampleCells[1, 1].PutValue("Completed");
                sampleCells[2, 0].PutValue("Task 2");
                sampleCells[2, 1].PutValue("Pending");
                sampleWb.Save(inputFile);
            }

            // Load the existing workbook (guard against missing file)
            if (!File.Exists(inputFile))
                throw new FileNotFoundException($"Input file '{inputFile}' not found.");

            var workbook = new Workbook(inputFile);
            var sheet = workbook.Worksheets[0];
            var cells = sheet.Cells;

            // Create a style with strikethrough font effect
            var strikeStyle = workbook.CreateStyle();
            strikeStyle.Font.IsStrikeout = true;

            // Assume column B (index 1) holds the task status
            int firstDataRow = 1; // row 0 contains headers
            int totalRows = cells.MaxDataRow;
            int totalCols = cells.MaxDataColumn;

            for (int row = firstDataRow; row <= totalRows; row++)
            {
                var statusCell = cells[row, 1];
                if (statusCell.StringValue.Equals("Completed", StringComparison.OrdinalIgnoreCase))
                {
                    // Apply the strikethrough style to every cell in the completed task row
                    for (int col = 0; col <= totalCols; col++)
                    {
                        cells[row, col].SetStyle(strikeStyle);
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputFile);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
