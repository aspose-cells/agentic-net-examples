using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaSignatureValidator
{
    public class VbaSignatureChecker
    {
        // Validates whether the VBA project inside the workbook referenced in the HTML is digitally signed.
        // The HTML is expected to contain a file path to an .xlsm workbook (e.g., <a href="C:\Files\sample.xlsm">).
        public static void ValidateVbaSignatureFromHtml(string htmlContent)
        {
            // Extract the first .xlsm file path from the HTML using a simple regular expression.
            // This is a lightweight approach; for complex HTML, consider using an HTML parser library.
            string pattern = @"href\s*=\s*[""'](?<path>[^""'>]+\.xlsm)[""']";
            Match match = Regex.Match(htmlContent, pattern, RegexOptions.IgnoreCase);
            if (!match.Success)
            {
                Console.WriteLine("No .xlsm workbook reference found in the provided HTML.");
                return;
            }

            string workbookPath = match.Groups["path"].Value;

            // Verify that the file exists before attempting to load it.
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook file not found at path: {workbookPath}");
                return;
            }

            try
            {
                // Load the workbook that potentially contains a VBA project.
                Workbook workbook = new Workbook(workbookPath);

                // Access the VBA project associated with the workbook.
                VbaProject vbaProject = workbook.VbaProject;

                // Check if a VBA project is present. Some workbooks may not contain VBA.
                if (vbaProject == null)
                {
                    Console.WriteLine("The workbook does not contain a VBA project.");
                    return;
                }

                // Determine whether the VBA code is signed.
                bool isSigned = vbaProject.IsSigned;
                Console.WriteLine($"VBA project signed: {isSigned}");

                // If signed, also verify the validity of the signature.
                if (isSigned)
                {
                    bool isValid = vbaProject.IsValidSigned;
                    Console.WriteLine($"VBA signature valid: {isValid}");
                }
            }
            catch (Exception ex)
            {
                // Handle any exceptions that may arise during loading or processing.
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
            }
        }

        // Example usage.
        public static void Main()
        {
            // Sample HTML containing a link to an .xlsm workbook.
            string sampleHtml = @"<html><body>
                <p>Download the macro-enabled workbook:</p>
                <a href=""C:\Temp\MacroWorkbook.xlsm"">MacroWorkbook.xlsm</a>
                </body></html>";

            ValidateVbaSignatureFromHtml(sampleHtml);
        }
    }
}