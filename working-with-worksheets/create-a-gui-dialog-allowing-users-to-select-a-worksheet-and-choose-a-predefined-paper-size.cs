// Title: Console application to select a worksheet and set its paper size using Aspose.Cells C# API
// Common Searches: how to set worksheet paper size with Aspose.Cells in a C# console program | C# code to list worksheets and let user choose one for page setup changes | enumerate PaperSizeType values and apply a selected size to an Excel sheet using Aspose.Cells | save an Excel workbook after modifying its PageSetup paper size in C# | prompt user for input and output file paths when changing worksheet page settings with Aspose.Cells
// Tags: Aspose.Cells set worksheet paper size | C# enumerate PaperSizeType enum | interactive worksheet selection Aspose.Cells | modify page setup and save workbook C# | console based Excel paper size selector

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPaperSizeSelector
{
    // Main entry point
    // The example loads an Excel workbook with Aspose.Cells, displays all worksheet names, asks the user to choose a worksheet by index, lists every PaperSizeType enum value, lets the user pick a paper size, applies the selected size to the chosen worksheet's PageSetup, and finally saves the updated workbook to a user‑specified location.
    static class Program
    {
        [STAThread]
        static void Main()
        {
            try
            {
                Console.WriteLine("=== Aspose.Cells Paper Size Selector ===");

                // Prompt for input workbook path
                Console.Write("Enter the full path of the Excel workbook to load: ");
                string inputPath = Console.ReadLine()?.Trim('"');

                if (string.IsNullOrWhiteSpace(inputPath) || !File.Exists(inputPath))
                {
                    Console.WriteLine("Error: File not found. Exiting.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Console.WriteLine($"Workbook loaded successfully. Contains {workbook.Worksheets.Count} worksheet(s).");

                // List worksheets
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    Console.WriteLine($"{i}: {workbook.Worksheets[i].Name}");
                }

                // Prompt for worksheet selection
                Console.Write("Select worksheet index: ");
                if (!int.TryParse(Console.ReadLine(), out int wsIndex) ||
                    wsIndex < 0 || wsIndex >= workbook.Worksheets.Count)
                {
                    Console.WriteLine("Invalid worksheet index. Exiting.");
                    return;
                }

                Worksheet ws = workbook.Worksheets[wsIndex];

                // List available paper sizes
                PaperSizeType[] paperSizes = (PaperSizeType[])Enum.GetValues(typeof(PaperSizeType));
                for (int i = 0; i < paperSizes.Length; i++)
                {
                    Console.WriteLine($"{i}: {paperSizes[i]}");
                }

                // Prompt for paper size selection
                Console.Write("Select paper size index: ");
                if (!int.TryParse(Console.ReadLine(), out int sizeIndex) ||
                    sizeIndex < 0 || sizeIndex >= paperSizes.Length)
                {
                    Console.WriteLine("Invalid paper size index. Exiting.");
                    return;
                }

                PaperSizeType selectedSize = paperSizes[sizeIndex];
                ws.PageSetup.PaperSize = selectedSize;
                Console.WriteLine($"Paper size set to {selectedSize} for worksheet \"{ws.Name}\".");

                // Prompt for output path
                Console.Write("Enter the full path to save the modified workbook (e.g., ModifiedWorkbook.xlsx): ");
                string outputPath = Console.ReadLine()?.Trim('"');

                if (string.IsNullOrWhiteSpace(outputPath))
                {
                    Console.WriteLine("Invalid output path. Exiting.");
                    return;
                }

                // Ensure directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
