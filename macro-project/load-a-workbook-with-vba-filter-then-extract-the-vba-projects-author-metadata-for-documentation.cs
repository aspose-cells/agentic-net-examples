using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaMetadata
{
    class Program
    {
        static void Main()
        {
            // Path to the macro‑enabled workbook to be loaded
            string inputPath = "sample_with_macro.xlsm";

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

                // Verify that the workbook actually contains a VBA project
                if (!workbook.HasMacro || workbook.VbaProject == null)
                {
                    Console.WriteLine("The specified workbook does not contain a VBA project.");
                    return;
                }

                // Access the VBA project
                VbaProject vbaProject = workbook.VbaProject;

                // Extract VBA project name (often used as the project identifier)
                string vbaProjectName = vbaProject.Name ?? "(Unnamed Project)";

                // Extract the author metadata from the built‑in document properties
                string author = "(Unknown Author)";
                var authorProp = workbook.BuiltInDocumentProperties["Author"];
                if (authorProp != null && authorProp.Value != null)
                {
                    author = authorProp.Value.ToString();
                }

                // Output the extracted information
                Console.WriteLine($"VBA Project Name: {vbaProjectName}");
                Console.WriteLine($"Workbook Author (metadata): {author}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}