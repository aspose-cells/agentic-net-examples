// Title: Set the default workbook style to a two-decimal-place percentage format using Aspose.Cells for .NET
// AI Prompts: Write C# code that changes a workbook's DefaultStyle to the custom number format "0.00%" so every cell displays percentages with two decimal places. | Show how to apply a workbook‑wide percentage format (0.00%) to a newly created Aspose.Cells workbook and save it as an .xlsx file.
// Common Searches: asp.net set default number format to percentage in Aspose.Cells workbook | c# apply 0.00% format to all cells Aspose.Cells | how to change default style to percentage format Aspose.Cells .NET | global percentage number format for new workbook using Aspose.Cells
// Tags: default workbook style custom format Aspose.Cells | apply workbook-wide percentage format Aspose.Cells | set default number format two decimal places .NET | percentage custom format 0.00% Excel Aspose.Cells | global cell style configuration Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // // Creates a new workbook, modifies its DefaultStyle to use the custom number format "0.00%" so every cell shows percentages with two decimal places, and saves the file as FormattedPercentages.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Retrieve the workbook's default style
                Style defaultStyle = workbook.DefaultStyle;

                // Set a custom number format to display percentages with two decimal places
                defaultStyle.Custom = "0.00%";

                // Apply the modified style back as the workbook's default style
                workbook.DefaultStyle = defaultStyle;

                // Define output file path
                string outputPath = "FormattedPercentages.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
