// Title: Copy VBA macro ‘InitializeReport’ between Excel workbooks with Aspose.Cells DesignerStorage (C#)
// Description: Shows how to load a macro‑enabled template, extract the designer storage of the VBA module ‘InitializeReport’ using GetDesignerStorage, create a new workbook, enable macros, add a designer‑type module, inject the stored binary with AddDesignerStorage, and save the result as an .xlsm file.
// Keywords: Aspose.Cells copy VBA macro | DesignerStorage GetDesignerStorage | AddDesignerStorage C# | EnableMacros Aspose.Cells | VbaModuleType Designer | macro‑enabled workbook .NET | copy VBA module programmatically
// Common Searches: Aspose.Cells copy VBA module from one workbook to another | GetDesignerStorage example C# | AddDesignerStorage Aspose.Cells tutorial | How to enable macros in a generated workbook using Aspose.Cells | Copy specific macro InitializeReport with Aspose.Cells
// Developer Intent: Transfer the ‘InitializeReport’ VBA macro from an existing .xlsm template to a newly created macro‑enabled workbook using Aspose.Cells DesignerStorage APIs in C#.
// Use Cases: Generate client‑specific reports that reuse a standard initialization macro. | Automate batch creation of macro‑enabled spreadsheets while sharing common utility macros. | Migrate legacy VBA code into dynamically generated workbooks without opening Excel. | Build a template‑driven workbook generator that reuses utility macros across projects.
// AI Prompts: Write C# code to copy multiple VBA modules from a template workbook to a target workbook using Aspose.Cells DesignerStorage. | Show how to verify a VBA module exists before copying and handle errors gracefully in Aspose.Cells. | Generate an example that copies a macro, updates its source code programmatically, and saves the modified workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace MacroCopyExample
{
    // Shows how to load a macro‑enabled template, extract the designer storage of the VBA module ‘InitializeReport’ using GetDesignerStorage, create a new workbook, enable macros, add a designer‑type module, inject the stored binary with AddDesignerStorage, and save the result as an .xlsm file.
    class Program
    {
        static void Main()
        {
            try
            {
                const string templatePath = "TemplateWithMacro.xlsm";
                const string targetPath = "TargetWithCopiedMacro.xlsm";

                // Verify that the template file exists
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {templatePath}");
                    return;
                }

                // Load the template workbook that contains the macro "InitializeReport"
                Workbook templateWorkbook = new Workbook(templatePath);

                // Retrieve the designer storage (binary data) of the macro module named "InitializeReport"
                byte[] macroStorage = templateWorkbook.VbaProject.Modules.GetDesignerStorage("InitializeReport");

                // Create an empty workbook for the target
                Workbook targetWorkbook = new Workbook();
                // Enable macros in the target workbook
                targetWorkbook.Settings.EnableMacros = true;

                // Add a designer module with the same name to the target workbook
                targetWorkbook.VbaProject.Modules.Add(VbaModuleType.Designer, "InitializeReport");

                // Insert the retrieved designer storage into the newly added module
                targetWorkbook.VbaProject.Modules.AddDesignerStorage("InitializeReport", macroStorage);

                // Save the target workbook as a macro‑enabled file
                targetWorkbook.Save(targetPath, SaveFormat.Xlsm);

                Console.WriteLine("Macro 'InitializeReport' copied successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
