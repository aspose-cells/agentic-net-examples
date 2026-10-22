// Title: Conditionally add opening password and protect workbook structure in Aspose.Cells for .NET when workbook has more than 10 worksheets
// Common Searches: Aspose.Cells C# set workbook password only if worksheet count exceeds 10 | How to protect Excel workbook structure conditionally with Aspose.Cells .NET | C# code to apply opening password to Excel file when it contains more than ten sheets using Aspose.Cells | Load password‑protected Excel file with fallback to unprotected load in Aspose.Cells | Check worksheet count before applying protection in Aspose.Cells workbook
// Tags: conditional workbook password protection Aspose.Cells | protect workbook structure Aspose.Cells C# | set opening password Aspose.Cells workbook | worksheet count based protection Aspose.Cells | load workbook with optional password Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file, handling both unprotected and password‑protected workbooks, checks if the workbook contains more than ten worksheets, and if so applies an opening password and protects the workbook structure before saving the file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        Workbook workbook = null;

        try
        {
            // Attempt to load the workbook without a password
            workbook = new Workbook(inputPath);
        }
        catch (CellsException ex) when (ex.Message.Contains("Please provide password"))
        {
            // If the workbook is password‑protected, provide the password here
            // Replace "YourPasswordHere" with the actual password if known
            var loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = "YourPasswordHere"
            };

            try
            {
                workbook = new Workbook(inputPath, loadOptions);
            }
            catch (Exception innerEx)
            {
                Console.WriteLine($"Failed to load password‑protected workbook: {innerEx.Message}");
                return;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        try
        {
            // If the workbook has more than 10 worksheets, apply protection
            if (workbook.Worksheets.Count > 10)
            {
                // Set opening password
                workbook.Settings.Password = "StrongPassword123";

                // Protect workbook structure (prevent adding/removing/renaming sheets)
                workbook.Protect(ProtectionType.Structure, "StrongPassword123");
            }

            // Save the workbook (overwrites or creates a new file)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during processing or saving: {ex.Message}");
        }
    }
}
