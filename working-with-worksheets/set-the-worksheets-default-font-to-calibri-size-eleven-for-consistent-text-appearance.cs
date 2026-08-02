// Title: Set Default Worksheet Font to Calibri 11 with Aspose.Cells for .NET
// Description: Learn how to change a workbook's DefaultStyle in Aspose.Cells for C# so that every new cell uses Calibri size 11. The example creates a workbook, updates the default font, writes a sample cell, and saves the file.
// Keywords: Aspose.Cells default font C# | set worksheet default style .NET | Calibri 11 Aspose.Cells | change workbook default font | Aspose.Cells default style font size | C# Excel default font Aspose
// Common Searches: Aspose.Cells set default font for workbook | C# change default worksheet font to Calibri | How to apply Calibri 11 as default style in Aspose.Cells | Set default font for all cells Aspose.Cells .NET | Aspose.Cells default style example
// Developer Intent: Configure the workbook’s default font to Calibri 11 so that all newly created cells inherit this style automatically.
// Use Cases: Generate reports that must follow a corporate Calibri 11 typography without setting the font on each cell. | Create templates where the default text appearance is predefined for consistency across multiple sheets. | Automate Excel file creation where the base font matches company branding standards.
// AI Prompts: Provide C# code using Aspose.Cells to set the workbook’s default font to Calibri 11 and save the document. | Explain how to modify the DefaultStyle of a workbook so new cells automatically use Calibri size 11. | Show how to revert the default font back to Arial 10 after it was changed to Calibri 11 with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Learn how to change a workbook's DefaultStyle in Aspose.Cells for C# so that every new cell uses Calibri size 11. The example creates a workbook, updates the default font, writes a sample cell, and saves the file.
    public class SetWorksheetDefaultFont
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Access the workbook's default style and modify its font
                Style defaultStyle = workbook.DefaultStyle;
                defaultStyle.Font.Name = "Calibri";
                defaultStyle.Font.Size = 11;

                // Apply the modified style back as the workbook's default style
                workbook.DefaultStyle = defaultStyle;

                // Optionally add some text to demonstrate the default font is applied
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("This text uses Calibri 11 as default font");

                // Save the workbook (lifecycle rule: save)
                workbook.Save("WorksheetDefaultFont.xlsx");
                Console.WriteLine("Workbook saved successfully as WorksheetDefaultFont.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetWorksheetDefaultFont.Run();
        }
    }
}
