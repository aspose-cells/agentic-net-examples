// Title: How to add a strikethrough font to rows with "Completed" status in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a strikeout font style with Aspose.Cells and applies it to every cell in rows where the Status column equals "Completed". | Write a C# program that loads an existing workbook, scans column B for the value "Completed", sets a strikethrough style on the entire matching rows, and saves the result to a new file. | Show how to define a reusable Aspose.Cells style with IsStrikeout = true and assign it to a range of rows based on a cell condition in C#.
// Common Searches: C# Aspose.Cells apply strikethrough to entire row based on cell value | How to format completed tasks with strikeout font in Excel using Aspose.Cells | Aspose.Cells set font.IsStrikeout for rows where column B = 'Completed' | Iterate rows in Aspose.Cells and change style conditionally in .NET | Save modified workbook after applying conditional row formatting with Aspose.Cells
// Tags: strikeout row formatting Aspose.Cells C# | conditional row style based on cell value Aspose.Cells | set font.IsStrikeout property Aspose.Cells | load workbook modify styles Aspose.Cells | save updated Excel file Aspose.Cells C#

using Aspose.Cells;
using System;
using System.IO;

// The program loads Tasks.xlsx, creates a style with Font.IsStrikeout = true, checks column B for "Completed", applies the style to every cell in matching rows, and saves the modified workbook as Tasks_Updated.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "Tasks.xlsx";
            const string outputPath = "Tasks_Updated.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Create a style with strikethrough font effect
            Style strikeStyle = workbook.CreateStyle();
            strikeStyle.Font.IsStrikeout = true;

            // Assume the "Status" column is column B (index 1) and data starts at row 2 (skip header)
            int firstDataRow = 1; // zero‑based index, so row 2 in Excel
            int totalRows = cells.MaxDataRow;
            int totalColumns = cells.MaxDataColumn;

            // Iterate through rows and apply the style to rows marked as "Completed"
            for (int row = firstDataRow; row <= totalRows; row++)
            {
                string status = cells[row, 1].StringValue?.Trim() ?? string.Empty;
                if (string.Equals(status, "Completed", StringComparison.OrdinalIgnoreCase))
                {
                    // Apply the strikethrough style to every cell in the current row
                    for (int col = 0; col <= totalColumns; col++)
                    {
                        cells[row, col].SetStyle(strikeStyle);
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
