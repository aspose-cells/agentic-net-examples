// Title: How to set a corporate default font for headings and body text in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set default workbook font to corporate typeface | Change default font for all cells in an Excel file with Aspose.Cells .NET | Apply custom corporate font to Excel workbook theme using Aspose.Cells API | Set default style font name when creating or loading a workbook with Aspose.Cells | Update Excel workbook theme to use a specific font for headings and body via Aspose.Cells
// Tags: Aspose.Cells set default style font | C# change workbook default font | apply corporate typeface to Excel workbook | modify Excel theme font Aspose.Cells | update default font name in Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample loads or creates an Excel workbook, assigns a corporate typeface to the workbook's DefaultStyle.Font.Name, and saves the file, illustrating how to define a default font for headings and body text with Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Determine input workbook path
                string inputPath = "input.xlsx";
                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new one
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook(); // creates a default workbook with one worksheet
                }

                // -----------------------------------------------------------------
                // Set a default font for the workbook.
                // Aspose.Cells does not expose a Theme class in all versions,
                // so we use the DefaultStyle to define the fallback font.
                // -----------------------------------------------------------------
                try
                {
                    // Apply the desired default font
                    workbook.DefaultStyle.Font.Name = "CorporateBodyFont";
                }
                catch (Exception ex)
                {
                    // Log any unexpected errors while modifying the default style
                    Console.WriteLine($"Warning: Unable to set default font – {ex.Message}");
                }

                // Save the updated workbook
                string outputPath = "UpdatedTheme.xlsx";
                try
                {
                    workbook.Save(outputPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error saving workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                // General error handling
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
