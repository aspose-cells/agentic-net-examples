// Title: Aspose.Cells for .NET – Set Workbook Default Number Format to Percentage (0.00%)
// Description: C# example that changes the workbook's DefaultStyle to the built‑in number format index 10 ("0.00%"), making every cell display percentages with two decimal places by default. Includes a verification cell and saves the file.
// Keywords: Aspose.Cells | .NET | C# | default style | percentage format | 0.00% format | built‑in format index 10 | set default number format | worksheet default style | apply style to all cells
// Common Searches: Aspose.Cells set default percentage format | C# default number format 0.00% Aspose | built‑in format index 10 percentage Aspose.Cells | apply default style to whole workbook .NET | how to make all cells show percentages with two decimals
// Developer Intent: Configure the workbook’s default style so every cell automatically uses a two‑decimal‑place percentage format.
// Use Cases: Create a new workbook where any inserted numeric value appears as a percentage with two decimals without extra formatting code. | Update an existing workbook to enforce consistent percentage display across all worksheets. | Validate the default format by writing a sample value to a cell and saving the workbook.
// AI Prompts: Generate C# code using Aspose.Cells to set the workbook default number format to "0.00%" for all cells. | Explain how to retrieve and apply the built‑in format index for a two‑decimal‑place percentage in Aspose.Cells .NET. | Show how to reapply the modified default style to existing worksheets after changing the number format.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // C# example that changes the workbook's DefaultStyle to the built‑in number format index 10 ("0.00%"), making every cell display percentages with two decimal places by default. Includes a verification cell and saves the file.
    public class SetWorksheetDefaultPercentageFormat
    {
        // Entry point for the console application
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (optional, just to demonstrate)
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the workbook's default style
            Style defaultStyle = workbook.DefaultStyle;

            // Set the built‑in number format to "0.00%" (percentage with two decimal places)
            // Built‑in format index 10 corresponds to "0.00%" as per Aspose.Cells documentation
            defaultStyle.Number = 10;

            // Apply the modified default style back to the workbook (the style object is a reference,
            // but assigning it back makes the intention explicit)
            workbook.DefaultStyle = defaultStyle;

            // Optionally, put a sample value to verify the format
            worksheet.Cells["A1"].PutValue(0.456); // 45.60%
            worksheet.Cells["A1"].SetStyle(defaultStyle);

            // Save the workbook
            string outputPath = "WorksheetDefaultPercentageFormat.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}
