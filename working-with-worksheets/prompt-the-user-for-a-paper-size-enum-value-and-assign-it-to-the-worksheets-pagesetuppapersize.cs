// Title: C# console program to set Aspose.Cells worksheet page size from user‑entered paper size
// Common Searches: c# console read paper size and apply to Aspose.Cells worksheet page setup | map user entered string like A4 or Letter to Aspose.Cells PaperSizeType enum | default to A4 when invalid paper size is provided in Aspose.Cells workbook | save Aspose.Cells workbook after setting page size with dynamic user input | ensure output folder exists before saving Aspose.Cells file in .NET
// Tags: set worksheet page size Aspose.Cells | parse paper size string to PaperSizeType | use A4 when paper size parsing fails | create output directory before workbook save | handle invalid enum parsing in Aspose.Cells | assign PaperSizeType to PageSetup in .NET

using System;
using System.IO;
using Aspose.Cells;

// The example prompts the user to enter a paper size (e.g., A4, Letter, Legal), converts common names or full enum strings to the Aspose.Cells PaperSizeType, falls back to PaperA4 for invalid input, creates a Workbook, assigns the resolved PaperSizeType to the first worksheet's PageSetup, ensures the output directory exists, and saves the workbook as output.xlsx while handling possible errors.
class Program
{
    static void Main()
    {
        try
        {
            // Prompt the user to enter a paper size (e.g., A4, Letter, Legal)
            Console.WriteLine("Enter paper size (e.g., A4, Letter, Legal):");
            string input = Console.ReadLine();

            // Treat a null or empty input as empty string
            if (string.IsNullOrWhiteSpace(input))
            {
                input = string.Empty;
            }

            // Try to map common names to Aspose.Cells PaperSizeType values
            PaperSizeType paperSize;
            bool parsed = false;

            // First, handle short names like "A4", "Letter", "Legal"
            switch (input.Trim().ToUpperInvariant())
            {
                case "A4":
                    paperSize = PaperSizeType.PaperA4;
                    parsed = true;
                    break;
                case "LETTER":
                    paperSize = PaperSizeType.PaperLetter;
                    parsed = true;
                    break;
                case "LEGAL":
                    paperSize = PaperSizeType.PaperLegal;
                    parsed = true;
                    break;
                default:
                    // Attempt to parse the full enum name (e.g., "PaperA4")
                    parsed = Enum.TryParse<PaperSizeType>(input, true, out paperSize);
                    break;
            }

            // If parsing still fails, default to A4
            if (!parsed)
            {
                Console.WriteLine("Invalid paper size entered. Defaulting to A4.");
                paperSize = PaperSizeType.PaperA4;
            }

            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Assign the selected paper size to the worksheet's PageSetup
            worksheet.PageSetup.PaperSize = paperSize;

            // Define output file path
            string outputPath = "output.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file (lifecycle rule: save)
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved with paper size: {paperSize}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
