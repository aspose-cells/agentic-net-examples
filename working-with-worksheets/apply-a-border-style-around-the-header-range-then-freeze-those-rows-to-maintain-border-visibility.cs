// Title: Add thick black borders to a header range and freeze the top row using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a style with thick black borders on all sides, applies it to cells A1:D1, and then freezes the first worksheet row using Aspose.Cells. | Show how to use Aspose.Cells StyleFlag to apply a border style to a header range and call FreezePanes so the header stays visible while scrolling. | Provide a complete .NET console example that writes header values, styles the header with thick borders, freezes the top row, and saves the workbook as Output.xlsx.
// Common Searches: Aspose.Cells C# apply thick border to header row and keep it visible with FreezePanes | how to freeze the first row after styling header cells in an Excel file using Aspose.Cells C# | example of creating a border style for A1:D1 and freezing panes in Aspose.Cells for .NET
// Tags: Aspose.Cells thick black border style | Aspose.Cells freeze top worksheet row | C# apply border to header range Aspose.Cells | Aspose.Cells StyleFlag all borders example | Aspose.Cells save workbook as xlsx | Aspose.Cells FreezePanes with header styling

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// // Creates a workbook, writes header values to A1:D1, applies a thick black border style to that range, freezes the first row, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Example header data
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["C1"].PutValue("Header3");
            sheet.Cells["D1"].PutValue("Header4");

            // Define the header range (A1:D1)
            AsposeRange headerRange = sheet.Cells.CreateRange("A1:D1");

            // Create a style with thick black borders on all sides
            Style borderStyle = workbook.CreateStyle();
            borderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thick;
            borderStyle.Borders[BorderType.TopBorder].Color = Color.Black;
            borderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
            borderStyle.Borders[BorderType.BottomBorder].Color = Color.Black;
            borderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thick;
            borderStyle.Borders[BorderType.LeftBorder].Color = Color.Black;
            borderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thick;
            borderStyle.Borders[BorderType.RightBorder].Color = Color.Black;

            // Apply the style to the header range
            StyleFlag flag = new StyleFlag { All = true };
            headerRange.ApplyStyle(borderStyle, flag);

            // Freeze the first row so the header stays visible while scrolling
            int totalRows = sheet.Cells.MaxDataRow + 1;
            int totalColumns = sheet.Cells.MaxDataColumn + 1;
            sheet.FreezePanes(1, 0, totalRows, totalColumns);

            // Determine output path and ensure directory exists
            string outputPath = "Output.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? Directory.GetCurrentDirectory();
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
