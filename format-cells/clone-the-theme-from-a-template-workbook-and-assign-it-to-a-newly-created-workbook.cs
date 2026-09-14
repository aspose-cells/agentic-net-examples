// Title: Clone an Excel workbook theme from a template file and assign it to a new workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Transfer the theme of a source .xlsx file to a freshly created Workbook object with Aspose.Cells CopyTheme in C#. | Create an empty workbook, load a template workbook, and programmatically copy its theme before saving the result using Aspose.Cells for .NET. | Implement error‑checked code that verifies a template file, copies its theme to a new workbook, and writes the output with Aspose.Cells.
// Common Searches: Aspose.Cells C# copy theme from template workbook to new workbook | how to use CopyTheme method in Aspose.Cells .NET example | programmatically duplicate an Excel theme using Aspose.Cells | apply template workbook theme to another file using Aspose.Cells C# | sample code for copying workbook theme in Aspose.Cells for .NET
// Tags: theme copy operation Aspose.Cells | clone Excel theme with Aspose.Cells | template theme assignment C# | transfer workbook theme Aspose.Cells | Aspose.Cells theme cloning example

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsThemeClone
{
    // The example loads a template Excel file, creates an empty workbook, copies the template's theme to the new workbook using the CopyTheme method, ensures the output directory exists, and saves the new workbook with the cloned theme.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string templatePath = "Template.xlsx";
                const string outputPath = "NewWorkbook.xlsx";

                // Verify that the template file exists to avoid FileNotFoundException
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {templatePath}");
                    return;
                }

                // Load the template workbook that contains the desired theme
                Workbook templateWorkbook = new Workbook(templatePath);

                // Create a new workbook (initially empty)
                Workbook newWorkbook = new Workbook();

                // Clone the theme from the template workbook to the new workbook
                // Use CopyTheme method (Theme property is read‑only)
                newWorkbook.CopyTheme(templateWorkbook);

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the new workbook with the cloned theme
                newWorkbook.Save(outputPath);
                Console.WriteLine($"New workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
