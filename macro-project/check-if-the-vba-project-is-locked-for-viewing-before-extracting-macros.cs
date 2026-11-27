// Title: Detect password‑protected VBA project before extracting macros with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# check if Excel VBA project is password protected | How to skip macro extraction when VBA project is locked using Aspose.Cells | Read VBA modules only if workbook.VbaProject.IsProtected is false in C# | Extract .bas files from an Excel workbook while handling protected VBA project with Aspose.Cells | Determine VBA project protection status before saving workbook in .NET
// Tags: Aspose.Cells VBA protection detection | Workbook.VbaProject.IsProtected usage | extract VBA modules to bas files C# | skip macro extraction for locked VBA project | save workbook after macro extraction Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;   // Namespace containing VbaModule and VbaProject

// The example loads an Excel workbook, uses Workbook.VbaProject.IsProtected to identify a password‑protected VBA project, and extracts each VbaModule's code to a .bas file only when the project is not locked, then saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Determine if the VBA project is password protected
            bool isProtected = false;
            if (workbook.VbaProject != null)
            {
                // Use the IsProtected property to detect protection
                isProtected = workbook.VbaProject.IsProtected;
            }

            if (isProtected)
            {
                Console.WriteLine("The VBA project is password protected. Cannot extract macros.");
            }
            else if (workbook.VbaProject != null && workbook.VbaProject.Modules != null)
            {
                // Extract macros from each VBA module
                foreach (VbaModule module in workbook.VbaProject.Modules)
                {
                    try
                    {
                        string code = module.Codes;
                        string fileName = $"{module.Name}.bas";

                        File.WriteAllText(fileName, code);
                        Console.WriteLine($"Extracted module: {module.Name} to \"{fileName}\"");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to extract module \"{module.Name}\": {ex.Message}");
                    }
                }
            }
            else
            {
                Console.WriteLine("No VBA project or modules found in the workbook.");
            }

            // Save the workbook (no modifications made)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
