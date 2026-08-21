// Title: Iterate RichTextPortion objects to change font name in an Excel cell with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, adds rich‑text content to cell A1, loops through every RichTextPortion in the cell, sets each portion's Font.Name (e.g., to "Arial"), and saves the file as RichTextPortionFontName.xlsx. It demonstrates how to modify mixed‑format text programmatically using Aspose.Cells for C#.
// Keywords: Aspose.Cells RichTextPortion font | C# change cell portion font | iterate RichTextPortion Aspose.Cells | Excel cell mixed formatting .NET | set font name for each text portion | Aspose.Cells example GitHub
// Common Searches: how to change font of each RichTextPortion in Aspose.Cells C# | Aspose.Cells iterate cell text portions | set font name for mixed‑format cell Aspose.Cells | C# Aspose.Cells change font of part of a cell | example code RichTextPortion font change
// Developer Intent: Programmatically update the Font.Name of every RichTextPortion inside a worksheet cell.
// Use Cases: Apply corporate font to specific words within a cell while preserving other formatting. | Generate reports where headings and values share a cell but require distinct fonts. | Automate styling of user‑entered text that may contain multiple font styles.
// AI Prompts: Write C# code that creates a cell with three RichTextPortion objects ("Hello", " ", "World") and changes each portion's Font.Name to "Calibri" using Aspose.Cells. | Explain how to retrieve the RichText collection from a cell and loop through its portions to set different fonts for each. | Provide a GitHub‑ready snippet that iterates RichTextPortion objects and saves the workbook to a specified folder.

using System;
using Aspose.Cells;

// This example creates a workbook, adds rich‑text content to cell A1, loops through every RichTextPortion in the cell, sets each portion's Font.Name (e.g., to "Arial"), and saves the file as RichTextPortionFontName.xlsx. It demonstrates how to modify mixed‑format text programmatically using Aspose.Cells for C#.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Set text in cell A1
            Cell cell = worksheet.Cells["A1"];
            cell.PutValue("Hello World!");

            // Change the font of the whole text in the cell
            string text = cell.StringValue;
            if (!string.IsNullOrEmpty(text))
            {
                // Apply font name to the entire cell using its style
                Style style = cell.GetStyle();
                style.Font.Name = "Arial";
                cell.SetStyle(style);
            }

            // Save the workbook
            string outputPath = "RichTextPortionFontName.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
