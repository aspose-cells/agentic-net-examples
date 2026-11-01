// Title: Change the modify (write‑protection) password of an existing Excel workbook with Aspose.Cells for .NET without affecting its data
// Common Searches: Aspose.Cells C# replace workbook write protection password while keeping content | how to set a new modify password on an existing .xlsx using Aspose.Cells .NET | programmatically update Excel file write‑protect password without re‑saving data in C# | change existing workbook's modify password with Aspose.Cells LoadOptions example | C# code to change Excel file's write‑protection password using Aspose.Cells API
// Tags: Aspose.Cells write protection password update | C# change Excel modify password | Aspose.Cells LoadOptions workbook loading | preserve workbook content during protection change | Aspose.Cells set new write‑protect password

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing XLSX file (creating a blank workbook if it does not exist), assigns a new write‑protection password via Workbook.Settings.WriteProtection.Password, and saves the file, leaving all worksheet data and formatting untouched.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the workbook
            string filePath = "input.xlsx";

            // Ensure the file exists; create an empty workbook if it does not.
            if (!File.Exists(filePath))
            {
                var newWb = new Workbook();
                newWb.Save(filePath, SaveFormat.Xlsx);
            }

            // Load options (set opening password if required)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            // loadOptions.Password = "oldOpenPassword"; // uncomment and set if needed

            // Load the workbook
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Set a new password for write‑protection
            workbook.Settings.WriteProtection.Password = "NewModifyPassword";

            // Save the workbook with the updated protection
            workbook.Save(filePath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
