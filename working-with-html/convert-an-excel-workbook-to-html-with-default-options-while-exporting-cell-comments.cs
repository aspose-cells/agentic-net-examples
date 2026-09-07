// Title: Convert an Excel workbook to HTML with default options while preserving cell comments using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads a .xlsx file, verifies the file exists, and saves it as an .html file with Aspose.Cells using the built‑in HtmlSaveOptions so that cell comments are exported. | Show how to wrap the Excel‑to‑HTML conversion in a try‑catch block and log success or error messages to the console.
// Common Searches: c# aspocells convert xlsx to html preserving comments | how to export Excel workbook to HTML with default settings using Aspose.Cells .NET | save workbook as html including cell notes Aspose.Cells example | Aspose.Cells HtmlSaveOptions default export comments | console application Excel to HTML conversion file existence check
// Tags: Aspose.Cells HTML export default settings | export Excel to HTML with comments C# | C# workbook to HTML conversion Aspose.Cells | cell comments preservation Aspose.Cells HTML | file existence validation before Aspose.Cells conversion

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies that 'input.xlsx' exists, loads it into an Aspose.Cells Workbook, creates a default HtmlSaveOptions instance, and saves the workbook as 'output.html', automatically including cell comments, with basic error handling and console output.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the Excel workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options (default options are sufficient for basic export)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Save the workbook as an HTML file with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
