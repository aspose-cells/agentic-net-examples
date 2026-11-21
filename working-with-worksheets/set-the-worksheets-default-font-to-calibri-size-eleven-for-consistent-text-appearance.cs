// Title: How to set the default worksheet font to Calibri 11pt using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set workbook default font to Calibri 11 | change default style font size in Excel file with Aspose.Cells .NET | set global font for all worksheets in a new workbook using Aspose.Cells | C# code to apply Calibri 11 as default font for Aspose.Cells workbook | Aspose.Cells default style font customization example
// Tags: default worksheet font configuration Aspose.Cells | modify workbook default style in C# | apply Calibri 11pt to new Excel workbook using Aspose.Cells | global font setting for Excel files .NET | set default font for all sheets Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new Workbook, retrieves its DefaultStyle, sets the font name to Calibri and size to 11 points, reassigns the modified style back to the workbook, ensures the output directory exists, saves the file as output.xlsx, and includes basic exception handling.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Modify the workbook's default style
                Style defaultStyle = workbook.DefaultStyle;
                defaultStyle.Font.Name = "Calibri";
                defaultStyle.Font.Size = 11;
                workbook.DefaultStyle = defaultStyle; // Apply changes

                // Define output path
                string outputPath = "output.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
