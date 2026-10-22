// Title: Export an Excel workbook to HTML with a 10‑second timeout in C# using Aspose.Cells
// Common Searches: c# aspose.cells export excel to html with a time limit | how to abort aspose.cells workbook.Save after a few seconds | aspose.cells html conversion cancellation example in C# | run aspose.cells save operation with timeout handling | prevent long‑running Excel to HTML export using Aspose.Cells in C#
// Tags: aspose.cells html export timeout | c# workbook save cancellation | excel to html conversion abort | aspose.cells interrupt monitor alternative | timeout control for aspose.cells save

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

// The example loads an existing Excel file with Aspose.Cells, sets up HtmlSaveOptions, and starts the workbook.Save operation on a separate task. It then waits up to 10 seconds for the export to finish; if the wait expires, the process is considered interrupted and a timeout message is shown. Successful completion and any exceptions are reported, demonstrating how to apply a timeout to an Aspose.Cells HTML export in C#.
class ExportToHtmlWithTimeout
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options (customize as needed)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Export the workbook to HTML with a timeout of 10 seconds
            Task exportTask = Task.Run(() => workbook.Save(outputPath, htmlOptions));

            if (!exportTask.Wait(TimeSpan.FromSeconds(10)))
            {
                // Timeout occurred
                Console.WriteLine("Export was interrupted due to timeout.");
            }
            else
            {
                Console.WriteLine("Export completed successfully.");
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
