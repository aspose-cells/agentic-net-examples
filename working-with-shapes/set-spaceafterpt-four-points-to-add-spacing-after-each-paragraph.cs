// Title: How to set a uniform row height of 20 points for all used rows in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells .NET set row height for every used row in Excel workbook | C# code to apply same row height to all rows based on used range with Aspose | How to programmatically set row height to 20 points in each worksheet using Aspose.Cells | Iterate worksheets and adjust row heights automatically in Aspose.Cells for .NET | Set uniform row height in existing Excel file with Aspose.Cells C#
// Tags: Aspose.Cells set row height programmatically | C# uniform Excel row height across worksheets | adjust row height based on used range Aspose | Excel row height automation .NET | set row height 20 points Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook (or creates a new one), loops through every worksheet, determines the last used row, and sets each row's height to 20 points before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Load existing workbook if it exists; otherwise create a new one
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the used range to determine the last row with data
                int maxRow = sheet.Cells.MaxDataRow;

                // Set each row's height to 20 points
                for (int row = 0; row <= maxRow; row++)
                {
                    sheet.Cells.SetRowHeight(row, 20);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
