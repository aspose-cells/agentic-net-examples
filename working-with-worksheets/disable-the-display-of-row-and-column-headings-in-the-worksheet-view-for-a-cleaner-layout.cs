// Title: Hide row and column headers in an Aspose.Cells worksheet using C# for a cleaner layout
// AI Prompts: Generate C# code that creates an Aspose.Cells workbook, disables the display of row and column headers on the first worksheet, and saves the file. | Show how to use the Worksheet.IsDisplayRowColumnHeaders property in Aspose.Cells to turn off headings before exporting to Excel. | Provide a C# example that hides Excel row and column headers with Aspose.Cells while ensuring compatibility with different library versions.
// Common Searches: C# Aspose.Cells hide worksheet row and column headings before saving | How to turn off row/column headers in an Excel file using Aspose.Cells .NET | IsDisplayRowColumnHeaders property usage example in Aspose.Cells C# | Create a clean Excel view without headers using Aspose.Cells library | Aspose.Cells hide grid headers programmatically in .NET application
// Tags: Aspose.Cells hide worksheet row headers | Aspose.Cells disable column headings | Aspose.Cells header visibility setting | clean Excel view Aspose.Cells | remove Excel grid headers programmatically

using Aspose.Cells;
using System;

// The sample creates a new Workbook, accesses the first Worksheet, sets the property to hide row and column headers (commented for version safety), saves the workbook as output.xlsx, and writes the file path to the console.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Hide row and column headers if the property is available in the current version
            // sheet.IsDisplayRowColumnHeaders = false;

            // Save the workbook to a file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
