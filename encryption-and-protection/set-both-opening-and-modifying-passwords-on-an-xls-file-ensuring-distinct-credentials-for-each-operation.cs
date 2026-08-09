// Title: Set distinct opening and modify passwords on an Excel workbook (XLSX) using Aspose.Cells for .NET
// Description: Creates a workbook, applies an encryption password (Settings.Password) and a separate write‑protection password (Settings.WriteProtection.Password), saves the file, reloads it with LoadOptions.Password, validates the edit password, updates a cell only when validation succeeds, and saves the modified version.
// Keywords: Aspose.Cells opening password | Aspose.Cells write protection | Excel file encryption .NET | set open and edit passwords Aspose | load password protected workbook C# | validate write protection password | protect XLSX Aspose.Cells
// Common Searches: how to set different open and edit passwords for Excel using Aspose.Cells | load and modify a password‑protected workbook with Aspose.Cells C# | validate write‑protection password before editing Excel file Aspose | Aspose.Cells example for workbook encryption and write protection
// Developer Intent: Apply separate opening and modifying passwords to an Excel workbook and verify edit rights before making changes.
// Use Cases: Secure a newly generated workbook with distinct passwords for opening and editing before distribution. | Open an encrypted workbook, confirm the write‑protection password, and perform authorized updates. | Automate validation of edit credentials in a protected Excel file to enforce data integrity.
// AI Prompts: Generate C# code that assigns an opening password and a different write‑protection password to an XLSX file using Aspose.Cells. | Show how to open a password‑protected workbook with Aspose.Cells, validate the modify password, update a cell only if the password is correct, and save the changes.

using System;
using System.IO;
using Aspose.Cells;

// Creates a workbook, applies an encryption password (Settings.Password) and a separate write‑protection password (Settings.WriteProtection.Password), saves the file, reloads it with LoadOptions.Password, validates the edit password, updates a cell only when validation succeeds, and saves the modified version.
class SetOpeningAndModifyingPasswords
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Password protected workbook");

            // Set the opening (encryption) password
            workbook.Settings.Password = "open123";

            // Set the modifying (write‑protection) password
            workbook.Settings.WriteProtection.Password = "modify456";

            // Save the workbook (XLSX format is used here)
            string filePath = "ProtectedWorkbook.xlsx";
            workbook.Save(filePath);

            // Ensure the file exists before attempting to load it
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"The file '{filePath}' was not found.");

            // Load the workbook using the opening password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = "open123"
            };
            Workbook loadedWorkbook = new Workbook(filePath, loadOptions);

            // Verify that the data can be read
            Console.WriteLine("Cell A1 value: " + loadedWorkbook.Worksheets[0].Cells["A1"].StringValue);

            // Validate the write‑protection password
            bool canModify = loadedWorkbook.Settings.WriteProtection.ValidatePassword("modify456");
            Console.WriteLine("Write‑protection password valid: " + canModify);

            // Example modification (only if password is correct)
            if (canModify)
            {
                loadedWorkbook.Worksheets[0].Cells["A2"].PutValue("Modified after password validation");
                // Save the modified workbook (still protected with the same passwords)
                loadedWorkbook.Save("ProtectedWorkbook_Modified.xlsx");
                Console.WriteLine("Modified workbook saved successfully.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
