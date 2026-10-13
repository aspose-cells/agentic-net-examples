// Title: Set the default worksheet font to Calibri 11 pt in a new Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that sets the workbook's DefaultStyle font to Calibri 11 pt before saving the file with Aspose.Cells. | Show how to configure a global font for all cells in a newly created Excel workbook using the Aspose.Cells .NET API. | Demonstrate applying a default cell style with Calibri 11 pt to a workbook and saving it as Output.xlsx.
// Common Searches: Aspose.Cells C# set default font for new workbook to Calibri 11 | how to change default worksheet style font size using Aspose.Cells .NET | global font setting for all cells in Aspose.Cells generated Excel file | C# Aspose.Cells default style font name and size example | apply Calibri 11pt as default font across worksheets with Aspose.Cells
// Tags: Aspose.Cells default font setting | C# workbook default style modification | apply Calibri 11pt to all cells | initialize Excel workbook with custom default style | Aspose.Cells global cell style configuration

using Aspose.Cells;
using System;
using System.IO;

// Creates a new Workbook, modifies its DefaultStyle to use Calibri 11 pt, ensures the output directory exists, and saves the workbook as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Retrieve the workbook's default style
            Style defaultStyle = workbook.DefaultStyle;

            // Set the default font to Calibri, size 11
            defaultStyle.Font.Name = "Calibri";
            defaultStyle.Font.Size = 11;

            // Apply the modified style back to the workbook
            workbook.DefaultStyle = defaultStyle;

            // Define output file path
            string outputPath = "Output.xlsx";

            // Ensure the directory exists
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
