// Title: Highlight all cells containing the word “error” (case‑insensitive) with an orange fill using Aspose.Cells for .NET
// AI Prompts: Search a worksheet for the text "error" without case sensitivity and set the cell background to orange using Aspose.Cells in C#. | Create a reusable method that takes a keyword and a color, finds every matching cell in an Excel file with Aspose.Cells, and applies the specified fill. | Extend the example to iterate through all worksheets in a workbook, highlighting case‑insensitive occurrences of a given term with a custom color.
// Common Searches: aspnet find cells containing a specific word case insensitive and change background color | c# aspose.cells highlight cells that contain 'error' orange | how to use FindOptions LookAtType.Contains to color matching cells in Excel with Aspose | search and format cells by text value using Aspose.Cells for .NET | apply conditional formatting programmatically with Aspose.Cells based on cell text
// Tags: Aspose.Cells FindOptions case‑insensitive text lookup | orange cell fill for matched keywords Aspose.Cells | programmatic highlighting of Excel cells by content C# | search and style .xlsx cells using Aspose.Cells API | bulk cell formatting based on keyword Aspose.Cells .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The program loads "input.xlsx", performs a case‑insensitive search for the word "error" in the first worksheet, applies an orange background to each matching cell, and saves the result as "output.xlsx".
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Configure find options: search values, contain the word, ignore case
                FindOptions findOptions = new FindOptions
                {
                    LookInType = LookInType.Values,
                    LookAtType = LookAtType.Contains,
                    CaseSensitive = false
                };

                // Find the first cell containing the word "error"
                Cell cell = sheet.Cells.Find("error", null, findOptions);

                // Loop through all matching cells
                while (cell != null)
                {
                    // Get the current style of the cell
                    Style style = cell.GetStyle();

                    // Set the fill color to orange
                    style.ForegroundColor = Color.Orange;
                    style.Pattern = BackgroundType.Solid;

                    // Apply the updated style back to the cell
                    cell.SetStyle(style);

                    // Find the next matching cell (continue search after the current cell)
                    cell = sheet.Cells.Find("error", cell, findOptions);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
