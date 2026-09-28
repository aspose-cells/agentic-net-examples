// Title: How to insert a manual horizontal page break after row 30 in an Excel worksheet with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an existing .xlsx file, adds a manual horizontal page break after row 30 on the first worksheet using Aspose.Cells, and saves the workbook. | Write a .NET console application that uses Aspose.Cells to place a page break before row 31 (zero‑based index 30) and verifies the break is added. | Create a snippet that demonstrates adding a horizontal page break to a worksheet at row 30 with Aspose.Cells and outputs the modified file path.
// Common Searches: Aspose.Cells add horizontal page break after row 30 C# example | C# set manual page break in Excel worksheet using Aspose.Cells | How to control pagination in an Excel file with Aspose.Cells .NET | Insert page break before row 31 Aspose.Cells API usage
// Tags: Aspose.Cells horizontal page break API | C# insert manual page break Excel worksheet | Excel pagination after specific row Aspose.Cells | Worksheet page break manipulation .NET | Aspose.Cells add page break to workbook

using System;
using System.IO;
using Aspose.Cells;

// Loads an existing workbook, inserts a manual horizontal page break after row 30 on the first worksheet using Aspose.Cells, and saves the updated file.
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
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index or name as needed)
            var worksheet = workbook.Worksheets[0];

            // Insert a manual horizontal page break after row 30
            // HorizontalPageBreaks.Add adds a break before the specified row (zero‑based index)
            // Adding a break before row 31 (index 30) creates a break after row 30
            worksheet.HorizontalPageBreaks.Add(30);

            // Save the workbook with the new page break
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log or display the exception details for debugging
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
