// Title: Add a clickable hyperlink with custom display text to a specific Excel cell using Aspose.Cells for .NET (C#)
// AI Prompts: Create C# code that uses Aspose.Cells to insert a hyperlink pointing to https://www.example.com into cell B2 and set the visible text to "Visit Example". | Show how to call HyperlinkCollection.Add for a single‑cell range and then assign a display string to the same cell in Aspose.Cells for .NET. | Write a reusable C# method that accepts a worksheet, cell address, link text, and URL, adds the hyperlink with the given text, and saves the workbook.
// Common Searches: aspnet add hyperlink with custom text to Excel cell using Aspose.Cells | c# Aspose.Cells HyperlinkCollection.Add example for single cell | how to set display text for a hyperlink in an Excel workbook with Aspose.Cells | insert clickable URL into specific cell B2 using Aspose.Cells .NET
// Tags: Aspose.Cells HyperlinkCollection.Add single-cell | C# set hyperlink display text Excel | Aspose.Cells insert clickable URL cell | save workbook after adding hyperlink Aspose.Cells | custom hyperlink text Excel Aspose.Cells

using System;
using Aspose.Cells;

namespace HyperlinkDemo
{
    // Demonstrates creating a new workbook, adding a hyperlink to cell B2 with the visible text "Visit Example", and saving the file as HyperlinkDemo.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Get the first worksheet
                var worksheet = workbook.Worksheets[0];

                // Target cell address for the hyperlink
                string cellAddress = "B2";

                // Convert cell address to zero‑based row and column indices
                int row = worksheet.Cells[cellAddress].Row;
                int column = worksheet.Cells[cellAddress].Column;

                // Add a hyperlink to the specified cell (single‑cell range)
                worksheet.Hyperlinks.Add(row, column, 1, 1, "https://www.example.com");

                // Set the display text that will be shown in the cell
                worksheet.Cells[cellAddress].PutValue("Visit Example");

                // Save the workbook
                string outputPath = "HyperlinkDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
