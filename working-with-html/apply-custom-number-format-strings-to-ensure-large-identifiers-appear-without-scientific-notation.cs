// Title: How to apply a custom number format to a cell range in Aspose.Cells for .NET to prevent scientific notation of large identifiers
// Common Searches: Aspose.Cells C# apply custom number format to prevent scientific notation | set number format "0" for a specific range of cells using Aspose.Cells .NET | avoid exponential display of large IDs in Excel with Aspose.Cells | use StyleFlag NumberFormat to format a column in Aspose.Cells C# | load workbook and format column A as plain numbers with Aspose.Cells
// Tags: custom number format string Aspose.Cells | cell range style application Aspose.Cells | disable scientific notation Excel .NET | StyleFlag NumberFormat property Aspose.Cells | format large ID column Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook, creates a style with the custom number format "0" to keep large identifiers from appearing in scientific notation, applies this style to the range A2:A100 using a StyleFlag that targets only the number format, and saves the updated workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file not found at '{inputPath}'.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Create a style with a custom number format to prevent scientific notation
            Style customStyle = workbook.CreateStyle();
            customStyle.Custom = "0";

            // Define the range containing large identifiers (e.g., A2:A100)
            Aspose.Cells.Range idRange = sheet.Cells.CreateRange("A2:A100");

            // Apply the custom style only to the number‑format property
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;
            idRange.ApplyStyle(customStyle, flag);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
