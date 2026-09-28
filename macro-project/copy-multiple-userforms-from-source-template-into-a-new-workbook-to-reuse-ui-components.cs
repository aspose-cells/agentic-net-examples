// Title: Copy VBA UserForms from a macro‑enabled Excel template to a new .xlsm workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a .xlsm template containing VBA UserForms with Aspose.Cells and saves it as a new .xlsm file while preserving all VBA components. | Show how to transfer only selected UserForms from a source workbook to a target workbook, keeping other VBA modules unchanged, using Aspose.Cells in C#. | Demonstrate merging UserForms from several macro‑enabled workbooks into a single .xlsm workbook with Aspose.Cells, ensuring VBA is retained.
// Common Searches: how to duplicate Excel UserForms programmatically with Aspose.Cells in C# | preserving VBA when saving a macro‑enabled workbook using Aspose.Cells .NET | copy specific UserForms from one .xlsm to another using Aspose.Cells | combine UserForms from multiple macro‑enabled Excel files with Aspose.Cells
// Tags: Aspose.Cells copy UserForms Xlsm | retain VBA modules Aspose.Cells .NET | selective UserForm export C# | merge macro-enabled workbooks Aspose.Cells | load template workbook retain VBA Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsVbaCopy
{
    // The example loads a macro‑enabled .xlsm template, verifies its existence, and saves a copy as a new .xlsm workbook using Aspose.Cells, preserving all VBA modules and UserForms while handling potential errors.
    class Program
    {
        static void Main(string[] args)
        {
            const string sourcePath = "SourceTemplate.xlsm";
            const string targetPath = "NewWorkbookWithUserForms.xlsm";

            try
            {
                // Verify that the source workbook exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the source workbook (macro‑enabled)
                var sourceWorkbook = new Workbook(sourcePath);

                // Save a copy of the workbook preserving VBA modules and UserForms
                sourceWorkbook.Save(targetPath, SaveFormat.Xlsm);

                Console.WriteLine($"Workbook saved successfully to '{targetPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
