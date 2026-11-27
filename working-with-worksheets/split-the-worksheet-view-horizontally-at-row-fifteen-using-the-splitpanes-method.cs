// Title: How to split a worksheet view horizontally at row 15 using Aspose.Cells FreezePanes in C#
// Common Searches: C# Aspose.Cells freeze panes after row 15 to create a horizontal split | How to create a horizontal split pane in Excel with Aspose.Cells API | Aspose.Cells FreezePanes example for splitting view at a specific row | Split worksheet view horizontally at row 15 using Aspose.Cells C# code | Freeze top rows in an Excel workbook with Aspose.Cells in C#
// Tags: Aspose.Cells FreezePanes horizontal split C# | split worksheet view at specific row Aspose.Cells | freeze top rows Excel Aspose.Cells C# | horizontal pane split Excel file using Aspose.Cells | C# create split panes after row 15 Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new workbook, fills 30 rows with data, and uses the FreezePanes method (15, 0, 15, 0) to split the worksheet view horizontally after row 15, then saves the file as SplitPanesDemo.xlsx.
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
            sheet.Name = "Demo";

            // Populate some data for visual verification
            for (int i = 0; i < 30; i++)
            {
                sheet.Cells[i, 0].PutValue($"Row {i + 1}");
                sheet.Cells[i, 1].PutValue(i * 10);
            }

            // Split the worksheet view horizontally after row 15
            // FreezePanes splits the view; the first three parameters specify the split position,
            // the last two specify the number of rows/columns to freeze in the top‑left pane.
            sheet.FreezePanes(15, 0, 15, 0);

            // Determine output path
            string outputPath = "SplitPanesDemo.xlsx";

            // Ensure the output directory exists
            string directory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
