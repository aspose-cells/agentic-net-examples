// Title: C# – Log Each Worksheet’s Automatic Paper Size with Aspose.Cells
// Description: Loads an Excel workbook using Aspose.Cells for .NET, iterates through all worksheets, reads the PageSetup.IsAutomaticPaperSize flag, and writes the sheet name, index and status to the console while handling missing files and runtime exceptions.
// Keywords: Aspose.Cells | C# workbook load | PageSetup.IsAutomaticPaperSize | worksheet automatic paper size | log worksheet page setup | iterate worksheets Aspose | Excel printing settings | console output C#
// Common Searches: Aspose.Cells read automatic paper size | C# get worksheet page setup properties | check IsAutomaticPaperSize for all sheets | log Excel sheet paper size using Aspose | print configuration audit Aspose.Cells
// Developer Intent: Load an existing Excel file and display whether each worksheet has the automatic paper‑size option enabled.
// Use Cases: Audit printing configuration before batch printing by confirming automatic paper size for every sheet. | Generate a compliance report of page‑setup flags across all worksheets. | Debug inconsistent page sizes in multi‑sheet workbooks. | Trigger conditional modifications to page setup based on the automatic‑size flag.
// AI Prompts: Write a method that disables PageSetup.IsAutomaticPaperSize for all worksheets in a loaded workbook using Aspose.Cells. | Show how to export the worksheet name, index, and automatic paper size flag to a CSV file instead of the console. | Provide enhanced error handling for workbook loading that distinguishes file‑not‑found, unsupported format, and permission‑denied scenarios.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an Excel workbook using Aspose.Cells for .NET, iterates through all worksheets, reads the PageSetup.IsAutomaticPaperSize flag, and writes the sheet name, index and status to the console while handling missing files and runtime exceptions.
    public class AutomaticPaperSizeLogger
    {
        public static void Run()
        {
            // Path to the input workbook
            string filePath = "input.xlsx";

            // Verify that the file exists before attempting to load it
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: The file '{filePath}' was not found.");
                return;
            }

            try
            {
                // Load the workbook from the specified file
                Workbook workbook = new Workbook(filePath);

                // Iterate through each worksheet and log its automatic paper size status
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    Worksheet sheet = workbook.Worksheets[i];
                    bool isAutomatic = sheet.PageSetup.IsAutomaticPaperSize;
                    Console.WriteLine($"Worksheet '{sheet.Name}' (Index {i}) - Automatic Paper Size: {isAutomatic}");
                }
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors during processing
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            AutomaticPaperSizeLogger.Run();
        }
    }
}
