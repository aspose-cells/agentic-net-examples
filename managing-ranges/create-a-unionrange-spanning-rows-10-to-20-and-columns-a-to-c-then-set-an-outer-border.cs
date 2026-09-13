// Title: Create a UnionRange for rows 10‑20 and columns A‑C and add a thick outer border using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that creates a UnionRange spanning rows 10 to 20 and columns A to C, then applies a thick border on all four sides of the range. | Show how to use Style and StyleFlag in Aspose.Cells to format the outer border of a specific cell range in an Excel workbook.
// Common Searches: asp.net aspose.cells how to define a union range covering rows 10 through 20 and columns A to C | apply thick outer border to a specific range in Excel using Aspose.Cells C# | set borders for a multi‑cell range with Aspose.Cells StyleFlag example | create and style a range A1:C11 in Aspose.Cells .NET | Aspose.Cells C# add border to UnionRange
// Tags: create union range Aspose.Cells | apply thick border Aspose.Cells | styleflag range formatting Aspose.Cells | excel range border formatting .NET | unionrange border styling C#

using Aspose.Cells;
using System;

// Creates a new workbook, defines a UnionRange covering rows 10‑20 and columns A‑C, applies a thick border on all sides using a Style and StyleFlag, saves the file as UnionRangeWithBorder.xlsx, and prints a success message.
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

            // Define a UnionRange covering rows 10 to 20 (indices 9‑19) and columns A to C (indices 0‑2)
            // Total rows = 11 (20‑10+1), total columns = 3
            Aspose.Cells.Range unionRange = sheet.Cells.CreateRange(9, 0, 11, 3);

            // Create a style for the outer border
            Style borderStyle = workbook.CreateStyle();
            borderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thick;
            borderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
            borderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thick;
            borderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thick;

            // Apply the style to the range
            StyleFlag flag = new StyleFlag { All = true };
            unionRange.ApplyStyle(borderStyle, flag);

            // Save the workbook
            string outputPath = "UnionRangeWithBorder.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
