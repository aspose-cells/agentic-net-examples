// Title: Highlight cells containing dd/MM/yyyy dates in column J (J1:J30) using a regular expression with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells find and highlight dates in dd/MM/yyyy format within a specific column range | How to apply a regex to cells J1:J30 in an Excel file using Aspose.Cells .NET | Using Aspose.Cells to search for date strings with pattern dd/MM/yyyy and set cell background color | Aspose.Cells C# example for regex-based date validation in column J | Highlight Excel cells that match a date pattern with Aspose.Cells and C#
// Tags: regex date search Aspose.Cells C# | highlight cells with Aspose.Cells style | search column J range J1:J30 Aspose.Cells | apply regex pattern dd/MM/yyyy Excel C# | conditional formatting using Aspose.Cells regex

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads or creates an Excel workbook, defines a regular expression for dates in the dd/MM/yyyy format, scans cells J1 through J30 for matches, collects the matching cells, applies a yellow background style to them via Aspose.Cells, and saves the modified workbook, handling any errors that may occur.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook or create a new one with sample data.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                for (int i = 0; i < 30; i++)
                {
                    ws.Cells[i, 9].PutValue("01/01/2023"); // Column J (index 9)
                }
            }

            Worksheet sheet = workbook.Worksheets[0];

            // Regex pattern for dates in dd/MM/yyyy format.
            string datePattern = @"\b(0[1-9]|[12][0-9]|3[01])/(0[1-9]|1[0-2])/\d{4}\b";
            Regex regex = new Regex(datePattern);

            // Find cells in J1:J30 that match the pattern.
            List<Cell> matchedCells = new List<Cell>();
            for (int row = 0; row < 30; row++)
            {
                Cell cell = sheet.Cells[row, 9]; // Column J
                if (regex.IsMatch(cell.StringValue))
                {
                    matchedCells.Add(cell);
                }
            }

            // Highlight matched cells with a yellow background.
            Style highlightStyle = workbook.CreateStyle();
            highlightStyle.ForegroundColor = Color.Yellow;
            highlightStyle.Pattern = BackgroundType.Solid;

            foreach (Cell cell in matchedCells)
            {
                cell.SetStyle(highlightStyle);
            }

            // Ensure the output directory exists.
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Processing complete. Output saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
