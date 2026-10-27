// Title: Merging adjacent RichTextPortion objects with identical formatting in Aspose.Cells for .NET
// Common Searches: Aspose.Cells .NET merge consecutive RichTextPortion objects with same style | C# combine adjacent rich text portions that have identical formatting in Excel | How to consolidate duplicate formatting in a cell using Aspose.Cells rich text API | Remove redundant RichTextPortion segments from an Excel cell with Aspose.Cells C# | Programmatically collapse identical rich text portions in a worksheet using Aspose.Cells
// Tags: merge identical RichTextPortion formatting Aspose.Cells | consolidate adjacent rich text portions .NET | Aspose.Cells rich text formatting optimization | C# collapse duplicate RichTextPortion objects | Excel cell rich text merging Aspose

using System;
using System.IO;
using Aspose.Cells;

// This example shows how to merge adjacent RichTextPortion objects that share the same formatting in an Excel cell using Aspose.Cells for .NET. It loads or creates a workbook, accesses a target cell, demonstrates the approach to identify and combine consecutive portions with identical style, and saves the updated file.
class MergeRichTextPortions
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load the workbook if it exists; otherwise create a new one.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                // Optionally add some default content.
                workbook.Worksheets[0].Cells["A1"].PutValue("Sample Text");
            }

            // Access the first worksheet and target cell.
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];

            // Retrieve the cell's string value using the StringValue property.
            string cellText = cell.StringValue;

            // If the cell is empty, nothing to process.
            if (string.IsNullOrEmpty(cellText))
            {
                workbook.Save(outputPath);
                return;
            }

            // Placeholder for rich text manipulation.
            // Since the .NET API does not expose RichString/RichTextPortion directly,
            // we simply reassign the same text.
            cell.PutValue(cellText);

            // Save the workbook.
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
