// Title: Prompt user for a paper size and set worksheet PageSetup.PaperSize with Aspose.Cells in C#
// AI Prompts: Write a C# program that reads a paper size name from the console, converts it to Aspose.Cells.PaperSizeType (case‑insensitive), assigns the enum value to the first worksheet's PageSetup.PaperSize, and saves the workbook. | Create a C# snippet that validates a user‑entered paper size string, maps it to the corresponding PaperSizeType enum, applies it to a worksheet's page setup, and gracefully handles invalid input.
// Common Searches: c# aspocells set worksheet paper size from user input | how to map console string to PaperSizeType enum in Aspose.Cells | changing page setup paper size programmatically with Aspose.Cells .NET | validate paper size entered by user before applying to workbook Aspose.Cells | example of setting A4 paper size using Aspose.Cells PageSetup
// Tags: set worksheet page setup paper size Aspose.Cells | parse console input to PaperSizeType enum | validate user‑provided paper size C# | apply paper size before saving workbook Aspose.Cells | handle invalid paper size input Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new Workbook, prompts the user to type a paper size (e.g., A4, Letter, Legal), parses the input into the Aspose.Cells.PaperSizeType enum with case‑insensitive matching, assigns the resulting enum to the first worksheet's PageSetup.PaperSize, and saves the file as Result.xlsx while handling empty or invalid entries and catching unexpected exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Prompt user for paper size enum value
            Console.WriteLine("Enter paper size (e.g., A4, Letter, Legal):");
            string input = Console.ReadLine();

            // Ensure input is not null or whitespace
            if (string.IsNullOrWhiteSpace(input))
            {
                Console.WriteLine("No input provided. No changes applied.");
            }
            else
            {
                // Parse the input to Aspose.Cells.PaperSizeType enum (case‑insensitive)
                if (Enum.TryParse<Aspose.Cells.PaperSizeType>(input, true, out Aspose.Cells.PaperSizeType paperSize))
                {
                    // Assign the parsed paper size to the worksheet's PageSetup
                    worksheet.PageSetup.PaperSize = paperSize;
                    Console.WriteLine($"Paper size set to {paperSize}.");
                }
                else
                {
                    Console.WriteLine("Invalid paper size entered. No changes applied.");
                }
            }

            // Save the workbook (lifecycle rule)
            workbook.Save("Result.xlsx");
            Console.WriteLine("Workbook saved as Result.xlsx");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
