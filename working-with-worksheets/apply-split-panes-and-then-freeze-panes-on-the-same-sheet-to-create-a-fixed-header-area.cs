// Title: Split panes and freeze top rows to create a fixed header in an Excel worksheet with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells that splits the worksheet at row 2 and freezes the first two rows as a permanent header. | Show how to call Worksheet.FreezePanes to create both a split and a frozen header area in an Excel file. | Adapt the example to also freeze the first column together with the top rows for a frozen side panel.
// Common Searches: aspnet c# how to split panes and freeze rows with Aspose.Cells | example of freezing top rows as header using Aspose.Cells Worksheet.FreezePanes | C# code to create split pane and fixed header in Excel workbook Aspose | freeze panes after splitting in Aspose.Cells .NET tutorial | set fixed header area in Excel using Aspose.Cells C#
// Tags: Aspose.Cells FreezePanes split pane | C# split pane Excel Aspose | fixed header Excel Aspose.Cells | Worksheet.FreezePanes top rows | Excel workbook freeze rows C#

using System;
using System.IO;
using Aspose.Cells;

// Creates a new workbook, fills 20 rows with sample data, applies FreezePanes(2,0,2,0) to split the sheet and freeze the first two rows as a fixed header, then saves the file as SplitAndFreeze.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data
            for (int i = 0; i < 20; i++)
            {
                sheet.Cells[i, 0].PutValue($"Row {i + 1}");
                sheet.Cells[i, 1].PutValue(i * 10);
            }

            // Freeze the top two rows (this also creates a split)
            sheet.FreezePanes(2, 0, 2, 0);

            // Save the workbook
            string outputPath = "SplitAndFreeze.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
