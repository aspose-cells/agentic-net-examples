// Title: How to insert a manual horizontal page break after row 30 in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Add a manual horizontal page break after row 30 to a worksheet loaded with Aspose.Cells and save the updated workbook. | Modify C# code to insert a horizontal page break at a configurable row index and confirm the pagination when printing the Excel file.
// Common Searches: Aspose.Cells C# add manual page break after row 30 | set horizontal page break programmatically with Aspose.Cells .NET | C# insert page break at specific row for printed Excel pagination using Aspose.Cells | example of using HorizontalPageBreaks collection in Aspose.Cells | how to control Excel print pagination via Aspose.Cells API
// Tags: Aspose.Cells horizontal page break API | insert manual page break C# Excel | control Excel pagination Aspose.Cells | add worksheet page break programmatically | horizontal page breaks collection Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, accesses the first worksheet, inserts a manual horizontal page break after row 30 via the HorizontalPageBreaks collection, and saves the modified file, with error handling for missing files and exceptions.
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
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a manual horizontal page break after row 30 (zero‑based index)
            sheet.HorizontalPageBreaks.Add(30);

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
