// Title: C# example: merge cells in a named range and add a thick border using Aspose.Cells
// Common Searches: C# Aspose.Cells merge cells in a named range and set thick borders | How to apply a thick border to a merged range with Aspose.Cells for .NET | Aspose.Cells code sample for merging a named range and formatting borders | Apply border style to merged cells in Excel using Aspose.Cells C# | Retrieve a named range, merge its cells, and add border styling with Aspose.Cells
// Tags: merge named range Aspose.Cells | apply thick border style Aspose.Cells | style merged cells Excel .NET | range border formatting C# | Aspose.Cells named range manipulation | Excel cell merging with border styling

using System;
using System.IO;
using Aspose.Cells;

// The program loads an Excel workbook, retrieves the named range "SummaryData", merges all its cells, creates a style with thick borders on every side, applies this style to the merged range, and saves the workbook as "output.xlsx".
class Program
{
    static void Main()
    {
        try
        {
            // Load the workbook (creates a new one if the file does not exist)
            Workbook workbook = LoadWorkbook("input.xlsx");

            // Retrieve the range defined by the named range "SummaryData"
            Aspose.Cells.Range summaryRange = workbook.Worksheets.GetRangeByName("SummaryData");
            if (summaryRange == null)
            {
                Console.WriteLine("Named range 'SummaryData' not found.");
                return;
            }

            // Merge all cells within the named range
            summaryRange.Merge();

            // Create a style with a thick border on all sides
            Style thickBorderStyle = workbook.CreateStyle();
            thickBorderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thick;
            thickBorderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
            thickBorderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thick;
            thickBorderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thick;

            // Apply the style to the merged range
            StyleFlag flag = new StyleFlag() { All = true };
            summaryRange.ApplyStyle(thickBorderStyle, flag);

            // Save the workbook (ensures the directory exists)
            SaveWorkbook(workbook, "output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Loads a workbook; creates a new one if the file is missing
    static Workbook LoadWorkbook(string path)
    {
        if (File.Exists(path))
        {
            return new Workbook(path);
        }
        else
        {
            // Create a new workbook with a default worksheet
            return new Workbook();
        }
    }

    // Saves a workbook; creates the directory if necessary
    static void SaveWorkbook(Workbook workbook, string path)
    {
        try
        {
            string directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }
            workbook.Save(path);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
            throw;
        }
    }
}
