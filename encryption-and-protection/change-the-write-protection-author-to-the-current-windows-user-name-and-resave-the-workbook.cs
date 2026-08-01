// Title: Set Excel Write‑Protection Author to Current Windows User with Aspose.Cells (C#)
// Description: Loads an existing .xlsx file using Aspose.Cells for .NET, sets Settings.WriteProtection.Author to the logged‑in Windows username, saves the workbook, and optionally reloads it to confirm the change.
// Keywords: Aspose.Cells | C# | write protection author | Excel metadata | Environment.UserName | update workbook author | set write protection | protect Excel file programmatically | save workbook Aspose | Excel security settings
// Common Searches: Aspose.Cells set write protection author C# | change Excel write protection author programmatically | assign current Windows user to Excel write protection author | update workbook metadata with Aspose.Cells | save workbook after modifying write protection
// Developer Intent: Assign the current Windows username to the workbook's write‑protection author and re‑save the file.
// Use Cases: Create audit‑ready Excel reports where the author field reflects the user who generated the file. | Automate metadata updates before distributing protected workbooks to ensure accurate author attribution. | Validate that write‑protection settings were applied correctly by reloading the file and reading the author property.
// AI Prompts: Generate C# code using Aspose.Cells to set the write‑protection author of an existing workbook to Environment.UserName and save it as a new file. | Explain the purpose of Settings.WriteProtection.Author in Aspose.Cells and show how to verify the value after saving. | Provide a concise example that loads a workbook, updates the write‑protection author with the current Windows user, saves the workbook, and prints the author to the console.

using System;
using Aspose.Cells;

namespace AsposeCellsWriteProtectionDemo
{
    // Loads an existing .xlsx file using Aspose.Cells for .NET, sets Settings.WriteProtection.Author to the logged‑in Windows username, saves the workbook, and optionally reloads it to confirm the change.
    class Program
    {
        static void Main()
        {
            // Path to the existing workbook
            string inputPath = "input.xlsx";

            // Load the workbook (create rule)
            Workbook workbook = new Workbook(inputPath);

            // Get the current Windows user name
            string currentUser = Environment.UserName;

            // Set the write‑protection author to the current user
            workbook.Settings.WriteProtection.Author = currentUser;

            // Save the workbook with the updated author (save rule)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);

            // Optional: verify the change
            Workbook verifyWorkbook = new Workbook(outputPath);
            Console.WriteLine("Write protection author: " + verifyWorkbook.Settings.WriteProtection.Author);
        }
    }
}
