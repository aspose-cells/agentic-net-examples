// Title: Set a worksheet to VeryHidden, apply FreezePanes, then restore visibility with Aspose.Cells for .NET (C#)
// Common Searches: how to set worksheet very hidden and freeze panes using Aspose.Cells C# | Aspose.Cells C# freeze panes on a hidden sheet before making it visible | C# code to hide a worksheet, apply FreezePanes at C3, and unhide with Aspose.Cells | using Aspose.Cells .NET to freeze rows and columns while worksheet is hidden | programmatically set IsVisible false, freeze panes, then set IsVisible true in Aspose.Cells
// Tags: Aspose.Cells very hidden worksheet handling | Aspose.Cells apply FreezePanes while sheet hidden | Aspose.Cells change IsVisible property C# | Aspose.Cells FreezePanes at C3 example | Aspose.Cells save workbook after visibility toggle

using System;
using System.IO;
using Aspose.Cells;

namespace WorksheetVisibilityAndFreezeDemo
{
    // The example creates a new workbook, fills it with sample data, sets the first worksheet to VeryHidden, freezes panes at cell C3 (freezing three rows and three columns) while the sheet remains hidden, then restores the worksheet's visibility and saves the file as an XLSX document using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Define output file path
                string outputPath = Path.Combine(Environment.CurrentDirectory, "WorksheetVisibilityAndFreezeDemo.xlsx");

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "DemoSheet";

                // Populate sample data to make the freeze effect visible
                sheet.Cells["A1"].PutValue("Header1");
                sheet.Cells["B1"].PutValue("Header2");
                sheet.Cells["C1"].PutValue("Header3");
                for (int row = 2; row <= 30; row++)
                {
                    sheet.Cells[$"A{row}"].PutValue($"Item {row - 1}");
                    sheet.Cells[$"B{row}"].PutValue(row * 10);
                    sheet.Cells[$"C{row}"].PutValue(DateTime.Now.AddDays(row - 2).ToShortDateString());
                }

                // 1. Hide the worksheet (Aspose.Cells uses IsVisible property; true = visible, false = hidden)
                sheet.IsVisible = false;
                Console.WriteLine($"Worksheet '{sheet.Name}' visibility set to hidden.");

                // 2. Freeze panes while the sheet is hidden
                // Freeze at cell C3 with 3 rows and 3 columns frozen
                sheet.FreezePanes("C3", 3, 3);
                Console.WriteLine("Panes frozen at C3 (3 rows, 3 columns).");

                // 3. Make the worksheet visible again
                sheet.IsVisible = true;
                Console.WriteLine($"Worksheet '{sheet.Name}' visibility restored to visible.");

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("An error occurred:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
