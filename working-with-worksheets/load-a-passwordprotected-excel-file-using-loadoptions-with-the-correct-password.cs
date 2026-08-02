// Title: Open a password‑protected Excel file with Aspose.Cells for .NET (C#) and save it unencrypted
// Description: Shows how to create a LoadOptions object with the workbook password, load the protected .xlsx using Aspose.Cells, clear the password setting, and write a new file without protection.
// Keywords: Aspose.Cells | C# | LoadOptions | password protected Excel | open encrypted workbook | remove Excel password | save unprotected workbook | .NET | Excel file encryption | Workbook.Settings.Password
// Common Searches: Aspose.Cells load encrypted Excel C# | How to open password protected .xlsx with Aspose.Cells | Remove password from Excel using Aspose.Cells .NET | LoadOptions password example Aspose.Cells | C# code to decrypt Excel workbook Aspose
// Developer Intent: Programmatically open a password‑protected Excel workbook using the correct password and then write it out without any protection.
// Use Cases: Read or modify data in a secured workbook | Batch‑process encrypted Excel files to create unprotected copies | Integrate password removal into automated data pipelines | Prepare password‑protected reports for downstream systems that cannot handle encryption
// AI Prompts: Generate C# code that opens a password‑protected .xlsx using Aspose.Cells LoadOptions and saves it without a password. | Write a script that scans a folder for .xlsx files, opens each with a given password via Aspose.Cells, removes the protection, and saves the result. | Explain how to use Workbook.Settings.Password to clear encryption after loading with LoadOptions.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to create a LoadOptions object with the workbook password, load the protected .xlsx using Aspose.Cells, clear the password setting, and write a new file without protection.
    public class LoadPasswordProtectedWorkbook
    {
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                // Path to the password‑protected Excel file
                string protectedFilePath = "protected.xlsx";

                // Verify that the source file exists
                if (!File.Exists(protectedFilePath))
                    throw new FileNotFoundException($"The file '{protectedFilePath}' was not found.");

                // Create LoadOptions and set the password required to open the workbook
                LoadOptions loadOptions = new LoadOptions
                {
                    Password = "test"
                };

                // Load the workbook using the LoadOptions with the correct password
                Workbook workbook = new Workbook(protectedFilePath, loadOptions);

                // Remove the password protection after loading
                workbook.Settings.Password = null;

                // Save the workbook without password protection
                string unprotectedFilePath = "unprotected.xlsx";
                workbook.Save(unprotectedFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
