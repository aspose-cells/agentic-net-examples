// Title: C# Aspose.Cells: Extract VBA Project Name and Workbook Author from an .xlsm File
// Description: Load a macro‑enabled workbook with Aspose.Cells for .NET, verify the VBA project, and retrieve its name, the workbook's Author property, and protection/signature status.
// Keywords: Aspose.Cells VBA project name | read workbook author property .NET | load xlsm with Aspose.Cells | VBA project protection check | extract VBA metadata C# | macro‑enabled Excel file Aspose | Aspose.Cells built‑in document properties | C# get VBA project details
// Common Searches: how to get VBA project name from xlsm using Aspose.Cells | retrieve author property from macro enabled workbook C# | check if VBA project is protected Aspose.Cells | read VBA project details programmatically | Aspose.Cells extract VBA metadata
// Developer Intent: Read a macro‑enabled Excel file and obtain VBA project details plus workbook author information.
// Use Cases: Document VBA projects by extracting names and author metadata for audit trails. | Automate compliance validation to ensure VBA projects are signed and not left unprotected before release. | Create an inventory of .xlsm files showing project name, author, and signature status for migration planning.
// AI Prompts: Generate C# code with Aspose.Cells that opens an .xlsm file and prints the VBA project name, workbook author, and whether the project is protected or signed. | Write a method using Aspose.Cells for .NET that returns true if the VBA project in a given workbook is protected. | Create a script that scans a folder of .xlsm files and logs each file's VBA project name, author metadata, and signed status.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaMetadata
{
    // Load a macro‑enabled workbook with Aspose.Cells for .NET, verify the VBA project, and retrieve its name, the workbook's Author property, and protection/signature status.
    class Program
    {
        static void Main()
        {
            // Path to the macro‑enabled workbook (xlsm) that contains VBA code
            string inputPath = "SampleWithVba.xlsm";

            // Verify that the file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"File not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Ensure the workbook actually contains a VBA project
                if (!workbook.HasMacro || workbook.VbaProject == null)
                {
                    Console.WriteLine("The specified workbook does not contain a VBA project.");
                    return;
                }

                // Access the VBA project
                VbaProject vbaProject = workbook.VbaProject;

                // Extract VBA project name
                string projectName = vbaProject.Name;
                Console.WriteLine($"VBA Project Name: {projectName}");

                // Extract author information from the workbook's built‑in document properties
                var authorProp = workbook.BuiltInDocumentProperties["Author"];
                if (authorProp != null && authorProp.Value != null)
                {
                    Console.WriteLine($"Workbook Author (metadata): {authorProp.Value}");
                }
                else
                {
                    Console.WriteLine("Author metadata not found in the workbook.");
                }

                // Additional optional information
                Console.WriteLine($"Is VBA Project Protected: {vbaProject.IsProtected}");
                Console.WriteLine($"Is VBA Project Signed: {vbaProject.IsSigned}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
