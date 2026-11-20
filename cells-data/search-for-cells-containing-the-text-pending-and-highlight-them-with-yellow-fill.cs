// Title: Highlight Excel cells containing the word "Pending" with a yellow fill using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells highlight cells that contain the word Pending | apply yellow background to cells with specific text in an Excel file using Aspose.Cells | search for string 'Pending' in all worksheets and set cell style with Aspose.Cells .NET | iterate through Excel cells and change fill color based on text using Aspose.Cells | load workbook, find cells containing Pending and save with highlighted cells Aspose.Cells C#
// Tags: highlight cells by text Aspose.Cells | apply solid yellow fill style C# | search string value in Excel worksheet Aspose.Cells | iterate all worksheets set cell style .NET | conditional cell formatting based on string Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The example loads an existing workbook (or creates a new one), creates a solid yellow fill style, iterates through every worksheet and cell, checks if a cell's string value contains "Pending", applies the yellow style to matching cells, and saves the modified workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        Workbook workbook = null;

        // Load workbook if it exists; otherwise create a new one
        if (File.Exists(inputPath))
        {
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (CellsException ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }
        }
        else
        {
            Console.WriteLine("Input file not found. Creating a new workbook.");
            workbook = new Workbook();
            workbook.Worksheets[0].Name = "Sheet1";
        }

        // Create a style with yellow fill
        Style yellowStyle = workbook.CreateStyle();
        yellowStyle.ForegroundColor = Color.Yellow;
        yellowStyle.Pattern = BackgroundType.Solid;

        // Iterate through all worksheets and apply style to cells containing "Pending"
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];
                    if (cell.Type == CellValueType.IsString && cell.StringValue.Contains("Pending"))
                    {
                        cell.SetStyle(yellowStyle);
                    }
                }
            }
        }

        // Save the modified workbook
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
