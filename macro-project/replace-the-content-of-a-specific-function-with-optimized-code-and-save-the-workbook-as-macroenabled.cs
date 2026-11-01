// Title: Replace a VBA function with optimized code in a macro-enabled Excel workbook and save as .xlsm using Aspose.Cells for .NET
// Common Searches: how to update a VBA function in an existing .xlsm file using Aspose.Cells C# | replace specific VBA module code with new implementation in Aspose.Cells .NET | save modified macro-enabled workbook as .xlsm after VBA code change with Aspose.Cells | using regex to replace a VBA function in an Aspose.Cells workbook C# example | optimize VBA routine in a macro-enabled Excel workbook programmatically with Aspose.Cells
// Tags: modify VBA module code Aspose.Cells | enhance VBA performance Aspose.Cells | save workbook as xlsm using Aspose.Cells | regex replace VBA code in Aspose.Cells | update VbaProject modules C#

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsMacroOptimization
{
    // The example loads a macro-enabled Excel file, locates a specific VBA function across all VBA modules, replaces the entire function definition with optimized code using a regular‑expression pattern, updates the module, and then saves the workbook as a new .xlsm file—all performed with Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "input.xlsm";
                const string outputPath = "output.xlsm";

                // Verify that the input workbook exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                    throw new FileNotFoundException($"Input file not found: {inputPath}");

                // Load the existing macro‑enabled workbook
                Workbook workbook = new Workbook(inputPath);

                // Optimize the VBA function named "MyFunction"
                OptimizeVbaFunction(workbook, "MyFunction", GetOptimizedFunctionCode());

                // Save the workbook as macro‑enabled (.xlsm)
                workbook.Save(outputPath, SaveFormat.Xlsm);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        /// <param name="workbook">The workbook containing the VBA project.</param>
        /// <param name="functionName">The name of the VBA function to replace.</param>
        /// <param name="optimizedCode">The full optimized function code (including Function/End Function).</param>
        private static void OptimizeVbaFunction(Workbook workbook, string functionName, string optimizedCode)
        {
            try
            {
                // Ensure the workbook has a VBA project
                if (workbook.VbaProject == null)
                    throw new InvalidOperationException("The workbook does not contain a VBA project.");

                // Iterate through all VBA modules to locate the function
                foreach (dynamic module in workbook.VbaProject.Modules)
                {
                    string originalCode = module.Codes as string ?? string.Empty;

                    // Build a regex pattern that matches the entire function definition (including its body)
                    string pattern = $@"Function\s+{Regex.Escape(functionName)}\b[^\r\n]*\r?\n([\s\S]*?)End\s+Function";

                    // Replace the matched function with the optimized version
                    string updatedCode = Regex.Replace(
                        originalCode,
                        pattern,
                        optimizedCode,
                        RegexOptions.IgnoreCase);

                    // If a replacement occurred, update the module's code and exit
                    if (!originalCode.Equals(updatedCode, StringComparison.Ordinal))
                    {
                        module.Codes = updatedCode;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error optimizing VBA function '{functionName}': {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Returns the optimized VBA function code.
        /// </summary>
        private static string GetOptimizedFunctionCode()
        {
            // Example of an optimized function; replace with real optimized logic.
            return
@"Function MyFunction()
    ' Optimized implementation starts here
    Dim result As Long
    result = 42  ' Placeholder for actual optimized calculation
    MyFunction = result
End Function";
        }
    }
}
