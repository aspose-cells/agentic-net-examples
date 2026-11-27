// Title: Set worksheet tab color to orange and freeze column A with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to change worksheet tab color programmatically | Freeze column A in Excel using Aspose.Cells FreezePanes method | Aspose.Cells example to set tab color and freeze panes in the same workbook | Apply orange tab color to first worksheet and freeze first column with Aspose.Cells | Save modified workbook as XLSX after changing tab color and freezing panes in C#
// Tags: worksheet tab color Aspose.Cells C# | freeze column A using FreezePanes Aspose.Cells | set orange tab color Excel Aspose.Cells | modify worksheet appearance Aspose.Cells .NET | save workbook as Xlsx Aspose.Cells C#

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The code loads an existing workbook or creates a new one, sets the first worksheet's tab color to orange, freezes column A using FreezePanes, and saves the workbook as output.xlsx.
class Program
{
    static void Main(string[] args)
    {
        try
        {
            // Load an existing workbook if a valid path is provided; otherwise create a new one.
            Workbook workbook;
            if (args.Length > 0 && File.Exists(args[0]))
            {
                string inputPath = args[0];
                workbook = new Workbook(inputPath);
                Console.WriteLine($"Loaded workbook from \"{inputPath}\".");
            }
            else
            {
                workbook = new Workbook();
                Console.WriteLine("Created a new workbook (no input file provided).");
            }

            // Access the first worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // -----------------------------------------------------------------
            // 1. Change the worksheet tab color.
            // -----------------------------------------------------------------
            worksheet.TabColor = Color.Orange;
            Console.WriteLine($"Worksheet tab color set to {worksheet.TabColor}.");

            // -----------------------------------------------------------------
            // 2. Freeze the first column (column A) so the colored tab stays visible.
            //    FreezePanes(row, column, freezedRows, freezedColumns)
            //    Row = 0 (no rows frozen), Column = 1 (split before column B),
            //    freezedRows = 0, freezedColumns = 1 (freeze one column).
            // -----------------------------------------------------------------
            worksheet.FreezePanes(0, 1, 0, 1);
            Console.WriteLine("First column frozen (column A).");

            // -----------------------------------------------------------------
            // 3. Save the modified workbook.
            // -----------------------------------------------------------------
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Report any unexpected errors.
            Console.Error.WriteLine($"Error: {ex.Message}");
            Console.Error.WriteLine(ex.StackTrace);
        }
    }
}
