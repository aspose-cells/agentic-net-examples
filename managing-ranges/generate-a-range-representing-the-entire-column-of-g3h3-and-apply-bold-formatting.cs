// Title: C# – Apply Bold Formatting to Entire Columns G and H with Aspose.Cells
// Description: Creates a new workbook, defines a range for G3:H3, retrieves the EntireColumn property to get columns G and H, builds a bold‑only style using StyleFlag, applies the style to the whole columns, and saves the file as BoldEntireColumns.xlsx.
// Keywords: Aspose.Cells C# | apply bold to column | EntireColumn property | range G3:H3 | StyleFlag font bold | format entire column Aspose.Cells | C# Excel formatting example | Aspose.Cells bold columns
// Common Searches: Aspose.Cells apply bold to whole column | C# format entire column G and H | How to use EntireColumn in Aspose.Cells | Bold specific columns in Excel with Aspose.Cells .NET | Create range and apply style flag Aspose.Cells
// Developer Intent: Apply a bold font style to every cell in the columns that intersect a given range.
// Use Cases: Highlight header columns in generated reports. | Enforce consistent bold styling for data columns based on a sample range. | Programmatically format selected columns without altering other cell attributes.
// AI Prompts: Show me C# code to bold entire columns G and H using Aspose.Cells. | How do I retrieve the EntireColumn of a range and apply only a bold font style? | Explain using StyleFlag to change just the font weight in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a new workbook, defines a range for G3:H3, retrieves the EntireColumn property to get columns G and H, builds a bold‑only style using StyleFlag, applies the style to the whole columns, and saves the file as BoldEntireColumns.xlsx.
    public class ApplyBoldToEntireColumns
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Create a range covering G3:H3 (row index 2, column index 6)
                Aspose.Cells.Range rangeG3H3 = cells.CreateRange(2, 6, 1, 2);

                // Get the entire columns that contain the range (columns G and H)
                Aspose.Cells.Range entireColumns = rangeG3H3.EntireColumn;

                // Create a style with bold font
                Style boldStyle = workbook.CreateStyle();
                boldStyle.Font.IsBold = true;

                // Apply only the bold attribute
                StyleFlag flag = new StyleFlag { FontBold = true };

                // Apply the bold style to the entire columns
                entireColumns.ApplyStyle(boldStyle, flag);

                // Save the workbook
                workbook.Save("BoldEntireColumns.xlsx");
                Console.WriteLine("Workbook saved as BoldEntireColumns.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the example
    public class Program
    {
        public static void Main()
        {
            ApplyBoldToEntireColumns.Run();
        }
    }
}
