// Title: Hide the Second Worksheet in an Excel Workbook with Aspose.Cells for .NET
// Description: Loads an existing Excel file, checks for a second worksheet, sets its IsVisible property to false to hide it, and saves the modified workbook to a new file using Aspose.Cells for C#.
// Keywords: Aspose.Cells hide worksheet C# | hide second sheet Aspose.Cells | C# hide Excel worksheet | Workbook.Save after hiding sheet | Aspose.Cells set IsVisible false | Excel sheet visibility .NET
// Common Searches: how to hide a specific worksheet using Aspose.Cells for .NET | C# code to hide the second sheet in an Excel file with Aspose.Cells | set worksheet visibility to hidden programmatically Aspose.Cells | Aspose.Cells hide worksheet and save workbook example | hide Excel sheet without opening Excel UI C#
// Developer Intent: Hide the second worksheet of an existing workbook and save the updated file.
// Use Cases: Publish a report where only the main sheet is visible while calculation sheets stay hidden. | Distribute a template that contains hidden helper sheets, exposing only the user‑editable sheet. | Prepare a workbook for external sharing by programmatically hiding auxiliary worksheets before saving.
// AI Prompts: Write C# code with Aspose.Cells to hide the third worksheet while leaving other sheets visible. | Show how to toggle worksheet visibility based on a runtime condition using Aspose.Cells for .NET. | Provide an example that loops through multiple worksheets, hides each one, and then saves the workbook.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an existing Excel file, checks for a second worksheet, sets its IsVisible property to false to hide it, and saves the modified workbook to a new file using Aspose.Cells for C#.
    public class HideSecondWorksheet
    {
        public static void Run()
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Hide the second worksheet if it exists
                if (workbook.Worksheets.Count > 1)
                {
                    workbook.Worksheets[1].IsVisible = false;
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            HideSecondWorksheet.Run();
        }
    }
}
