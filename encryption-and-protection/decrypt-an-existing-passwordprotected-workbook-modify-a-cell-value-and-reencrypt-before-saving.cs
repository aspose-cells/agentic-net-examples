// Title: Open, edit, and re‑encrypt a password‑protected Excel workbook with Aspose.Cells for .NET (C#)
// Description: Shows how to load an encrypted .xlsx using LoadOptions.Password, change the value of cell A1, set a new workbook password via Workbook.Settings.Password, and save the file with the updated encryption.
// Keywords: Aspose.Cells load encrypted workbook C# | modify cell in protected Excel file | change workbook password Aspose.Cells | re‑encrypt Excel workbook .NET | C# password protected Excel example | LoadOptions.Password Aspose | Workbook.Settings.Password usage
// Common Searches: Aspose.Cells open password protected Excel and edit cell | C# change password of encrypted workbook with Aspose | How to re‑encrypt an Excel file after modification using Aspose.Cells | LoadOptions.Password example for .xlsx in .NET | Update cell in protected workbook and save with new password
// Developer Intent: The developer needs to programmatically open a secured Excel file, modify its contents, and save it again under a different password.
// Use Cases: Automate the refresh of confidential report headers while rotating the protection password for each distribution cycle. | Batch‑process a library of encrypted workbooks to enforce a corporate password policy without manual steps. | Expose a web API that accepts a protected workbook, corrects data errors, and returns the file re‑encrypted with a new password.
// AI Prompts: Write C# code with Aspose.Cells that opens an encrypted workbook, updates cell B2, and saves it using a new password. | Explain the steps to change a workbook's password in Aspose.Cells without affecting formatting or other protection settings. | Provide robust error‑handling for loading a password‑protected Excel file and saving it with a different password using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsPasswordDemo
{
    // Shows how to load an encrypted .xlsx using LoadOptions.Password, change the value of cell A1, set a new workbook password via Workbook.Settings.Password, and save the file with the updated encryption.
    class Program
    {
        static void Main()
        {
            // Path to the existing password‑protected workbook
            string inputFile = "protected.xlsx";

            // Password required to open the workbook
            string openPassword = "oldpwd";

            // New password to protect the workbook after modification
            string newPassword = "newpwd";

            // Path for the modified and re‑encrypted workbook
            string outputFile = "modified_protected.xlsx";

            // Load the workbook using LoadOptions with the opening password
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = openPassword;
            Workbook workbook = new Workbook(inputFile, loadOptions);

            // Modify a cell value (e.g., set A1 to a new string)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Modified value");

            // Re‑apply password protection (encrypt the workbook)
            workbook.Settings.Password = newPassword;

            // Save the workbook; it will be saved with the new password
            workbook.Save(outputFile);

            // Optional: clean up
            workbook.Dispose();

            Console.WriteLine("Workbook decrypted, modified, and re‑encrypted successfully.");
        }
    }
}
