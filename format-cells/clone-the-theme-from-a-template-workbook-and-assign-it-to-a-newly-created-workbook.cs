// Title: Copy an Excel workbook theme from a template file to a new workbook with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# copy theme from existing workbook to new workbook | How to transfer Excel theme using Aspose.Cells .NET API | Clone workbook theme in C# with Aspose.Cells CopyTheme method | Apply template workbook theme to a newly created Excel file using Aspose.Cells
// Tags: CopyTheme Aspose.Cells C# | clone Excel theme using Aspose.Cells | transfer workbook theme .NET | apply template theme to new workbook | Aspose.Cells theme copying

using System;
using System.IO;
using Aspose.Cells;

// The example loads a template Excel file, creates an empty workbook, copies the theme from the template to the new workbook using the CopyTheme method, and saves the resulting workbook with the cloned theme.
class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "Template.xlsx";
            const string outputPath = "NewWorkbook.xlsx";

            // Verify that the template file exists before loading
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            // Load the template workbook that contains the desired theme
            Workbook templateWorkbook = new Workbook(templatePath);

            // Create a new workbook (initially empty)
            Workbook newWorkbook = new Workbook();

            // Copy the theme from the template workbook to the new workbook
            newWorkbook.CopyTheme(templateWorkbook);

            // Save the new workbook with the cloned theme
            newWorkbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
