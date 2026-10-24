// Title: Set worksheet page setup paper size from console input using Aspose.Cells for .NET
// Common Searches: how to change Excel worksheet paper size based on user input with Aspose.Cells C# | C# console program to set Aspose.Cells worksheet PageSetup PaperSize enum | validate PaperSizeType enum value before assigning to worksheet PageSetup in Aspose.Cells | prompt for numeric paper size code and apply to Aspose.Cells workbook | set custom paper size for XLSX file using Aspose.Cells .NET
// Tags: worksheet page setup paper size Aspose.Cells | enum validation for PaperSizeType C# | assign PaperSizeType to PageSetup in Aspose.Cells | console-driven paper size selection Aspose.Cells | save workbook to XLSX with Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a new Workbook, reads a numeric value from the console, validates it against the PaperSizeType enum, assigns the selected enum to the worksheet's PageSetup.PaperSize, and saves the workbook as an XLSX file, handling any errors that may occur.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Prompt the user to enter a PaperSizeType enum value (numeric)
            Console.WriteLine("Enter a paper size enum value (e.g., 9 for A4, 1 for Letter):");
            string input = Console.ReadLine() ?? string.Empty;

            // Try to convert the input to an integer and validate it against the PaperSizeType enum
            if (int.TryParse(input, out int sizeValue) &&
                Enum.IsDefined(typeof(PaperSizeType), sizeValue))
            {
                // Assign the selected paper size to the worksheet's PageSetup
                sheet.PageSetup.PaperSize = (PaperSizeType)sizeValue;
                Console.WriteLine($"Paper size set to {sheet.PageSetup.PaperSize}.");
            }
            else
            {
                Console.WriteLine("Invalid input. Using the default paper size.");
            }

            // Save the workbook to a file
            workbook.Save("output.xlsx");
            Console.WriteLine("Workbook saved as output.xlsx");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
