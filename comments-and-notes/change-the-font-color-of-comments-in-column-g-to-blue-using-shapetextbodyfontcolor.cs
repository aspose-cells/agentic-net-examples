// Title: How to change Excel comment font color to blue in column G using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# change comment text color in a specific column | set blue font color for cell comments in column G with Aspose.Cells | C# code to update Excel comment color using Aspose.Cells API | modify comment appearance in worksheet column G using Aspose.Cells for .NET | change comment font color programmatically in Excel with Aspose.Cells C#
// Tags: Aspose.Cells change comment font color | Excel comment color modification C# | column G comment styling Aspose.Cells | set comment text color blue Aspose.Cells | programmatic Excel comment formatting .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, iterates through every cell in column G, checks for existing comments, changes each comment's font color to blue, and saves the updated file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Column G (zero‑based index 6)
            const int columnIndex = 6;

            // Determine the last used row in the sheet
            int maxRow = sheet.Cells.MaxDataRow;

            // Iterate through each cell in column G
            for (int row = 0; row <= maxRow; row++)
            {
                Cell cell = sheet.Cells[row, columnIndex];

                // If the cell has a comment, change its font color to blue
                if (cell.Comment != null)
                {
                    cell.Comment.Font.Color = Color.Blue;
                }
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
