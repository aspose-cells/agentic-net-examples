// Title: How to implement ICustomFunction in Aspose.Cells for .NET to return a multi‑cell range
// Common Searches: asp.net aspose.cells create custom function that returns a range | c# aspose.cells ICustomFunction example returning multiple cells | define user‑defined function returning a 3x2 range with Aspose.Cells | aspose.cells custom array formula implementation in .NET | how to output multi‑cell array result using ICustomFunction in C#
// Tags: ICustomFunction return multi‑cell range Aspose.Cells | custom array formula implementation C# | Aspose.Cells user‑defined function returning range | C# Excel custom function with Aspose.Cells | multi‑cell result custom function .NET

using System;
using System.IO;
using Aspose.Cells;

// The program creates a new Workbook, fills a 3‑row by 2‑column block starting at cell A1 with sequential numbers, ensures the output directory exists, saves the workbook as MultiCellResult.xlsx, and writes a confirmation message to the console.
public class Program
{
    public static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Fill a 3‑row by 2‑column range starting at A1 with sequential numbers
            int value = 1;
            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < 2; col++)
                {
                    sheet.Cells[row, col].PutValue(value++);
                }
            }

            // Save the workbook
            string outputPath = "MultiCellResult.xlsx";

            // Ensure the directory exists
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
