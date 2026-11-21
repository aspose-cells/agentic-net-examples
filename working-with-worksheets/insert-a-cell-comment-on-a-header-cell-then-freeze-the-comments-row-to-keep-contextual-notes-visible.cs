// Title: Insert a persistent comment on a header cell and freeze the top row with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add always-visible comment to cell A1 | How to freeze the first worksheet row while keeping a comment displayed using Aspose.Cells | Aspose.Cells example for adding a comment and applying FreezePanes in C# | Make Excel header comment stay visible after freezing panes with Aspose.Cells | Persist comment visibility when using FreezePanes in Aspose.Cells for .NET
// Tags: add persistent comment Aspose.Cells C# | freeze first row Aspose.Cells | always visible comment Excel Aspose.Cells | header cell annotation Aspose.Cells | comment visibility with FreezePanes Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, writes "Header" to cell A1, adds an always-visible comment to that cell, freezes the first worksheet row with FreezePanes, and saves the file as CommentAndFreeze.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add header text to cell A1
            sheet.Cells["A1"].PutValue("Header");

            // Insert a comment on the header cell (A1)
            int commentIndex = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "This column contains important data.";
            comment.IsVisible = true; // Keep the comment always visible

            // Freeze the first row (row index 1) while keeping all columns unfrozen
            // Parameters: row, column, totalRowsToFreeze, totalColumnsToFreeze
            sheet.FreezePanes(1, 0, 1, 0);

            // Determine output path
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CommentAndFreeze.xlsx");

            // Ensure the directory exists
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
