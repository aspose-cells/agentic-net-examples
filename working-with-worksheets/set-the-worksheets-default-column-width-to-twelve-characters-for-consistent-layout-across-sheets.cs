// Title: Set Default Column Width (12 characters) for a Worksheet using Aspose.Cells for .NET (C#)
// Description: Creates a new Workbook, accesses the first Worksheet, sets Cells.StandardWidth to 12 characters, verifies the value, and saves the file as DefaultColumnWidth.xlsx with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | default column width | StandardWidth | worksheet column width | set column width 12 characters | Excel automation | Aspose.Cells API | column width all columns
// Common Searches: Aspose.Cells set default column width C# | How to use Cells.StandardWidth property | Set worksheet column width to 12 characters in .NET | Change default column width for new workbook Aspose.Cells | Excel column width automation with Aspose.Cells
// Developer Intent: Define a uniform default column width of 12 characters for a worksheet so every column inherits this size.
// Use Cases: Generate a template workbook where all columns start with a consistent width before data entry. | Create reports that require fixed column widths for alignment and readability across multiple sheets. | Automate Excel file creation where column layout must match corporate style guidelines.
// AI Prompts: Show C# code to set the default column width for every worksheet in an existing workbook using Aspose.Cells. | Provide an example that changes the default column width to 15 characters and then autosizes specific columns. | Explain the measurement unit of Cells.StandardWidth and how it affects column rendering in Excel.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a new Workbook, accesses the first Worksheet, sets Cells.StandardWidth to 12 characters, verifies the value, and saves the file as DefaultColumnWidth.xlsx with Aspose.Cells for .NET.
    public class SetDefaultColumnWidth
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Set the default column width to 12 characters
                worksheet.Cells.StandardWidth = 12.0;

                // Optional: verify the setting
                Console.WriteLine("Default column width set to: " + worksheet.Cells.StandardWidth + " characters");

                // Save the workbook
                string outputPath = "DefaultColumnWidth.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SetDefaultColumnWidth.Run();
        }
    }
}
