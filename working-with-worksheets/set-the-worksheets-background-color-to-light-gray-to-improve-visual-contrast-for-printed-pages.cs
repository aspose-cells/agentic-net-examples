// Title: How to set an Excel worksheet tab color to light gray using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that changes the first worksheet's tab color to LightGray and saves the workbook to a specified file path. | Provide a C# snippet that applies a custom RGB tab color to a worksheet, creates the output folder if it does not exist, and writes the file using Aspose.Cells.
// Common Searches: example of setting worksheet tab color to light gray with Aspose.Cells in C# | how to change Excel sheet tab color programmatically using Aspose.Cells .NET | C# create missing directory before saving Excel workbook with Aspose.Cells | customizing worksheet tab appearance for printed Excel files using Aspose.Cells | Aspose.Cells set worksheet tab color for better print contrast
// Tags: Aspose.Cells worksheet tab color setting | C# worksheet tab LightGray | create output folder Aspose.Cells | save workbook to path Aspose.Cells | RGB worksheet tab color Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The example creates a new Workbook, accesses the first Worksheet, sets its TabColor to LightGray, ensures the output directory exists, and saves the workbook as output.xlsx, handling any exceptions that may occur.
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

            // Set the worksheet tab color (as a visual indicator)
            sheet.TabColor = Color.LightGray;

            // Define output file path
            string outputPath = "output.xlsx";

            // Ensure the directory exists (if a directory part is present)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? string.Empty;
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
