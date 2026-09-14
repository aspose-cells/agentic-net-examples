// Title: Set Calibri 11pt as the default font for all worksheets using Aspose.Cells for .NET
// AI Prompts: Create a new Workbook, build a Style with Font.Name = "Calibri" and Font.Size = 11, assign it to Workbook.DefaultStyle, then save the workbook. | Open an existing Excel file, modify its Workbook.DefaultStyle to use Calibri 11pt, and overwrite the file with the updated style.
// Common Searches: Aspose.Cells .NET set default worksheet font to Calibri 11 | C# apply a global font style to every sheet in a new Excel workbook with Aspose | How to change the default font for all worksheets in an Aspose.Cells workbook | Set workbook.DefaultStyle font name and size using Aspose.Cells C#
// Tags: Aspose.Cells workbook default style configuration | global font Calibri 11pt for .NET Excel workbook | C# set workbook default font Aspose.Cells | apply default style to all sheets Aspose.Cells | initialize workbook with specific font Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, defines a Calibri 11pt style, assigns it to the workbook's DefaultStyle (affecting every worksheet), ensures the output directory exists, and saves the file as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Create a style with Calibri font, size 11
            Style defaultStyle = workbook.CreateStyle();
            defaultStyle.Font.Name = "Calibri";
            defaultStyle.Font.Size = 11;

            // Set this style as the workbook's default style (applies to all worksheets)
            workbook.DefaultStyle = defaultStyle;

            // Define output file path
            string outputPath = "Result.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
