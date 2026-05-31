using System;
using Aspose.Cells;

namespace AsposeCellsPaperSizeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the create rule)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Prompt the user to enter a paper size value.
            // The user should enter the numeric value of the PaperSizeType enum.
            Console.WriteLine("Enter the numeric value for the desired paper size (e.g., 9 for PaperA4):");
            string input = Console.ReadLine();

            // Try to parse the input to an integer.
            if (int.TryParse(input, out int paperSizeValue) &&
                Enum.IsDefined(typeof(PaperSizeType), paperSizeValue))
            {
                // Convert the integer to the corresponding enum value.
                PaperSizeType selectedSize = (PaperSizeType)paperSizeValue;

                // Assign the selected paper size to the worksheet's PageSetup.
                worksheet.PageSetup.PaperSize = selectedSize;

                Console.WriteLine($"Paper size set to: {selectedSize}");
            }
            else
            {
                Console.WriteLine("Invalid input. Using default paper size (PaperLetter).");
                worksheet.PageSetup.PaperSize = PaperSizeType.PaperLetter;
            }

            // Save the workbook (uses the save rule)
            string outputPath = "PaperSizeDemo.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}