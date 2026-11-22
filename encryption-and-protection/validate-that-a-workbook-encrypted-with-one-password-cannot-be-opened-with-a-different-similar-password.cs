// Title: Confirm that an Aspose.Cells encrypted XLSX cannot be opened with an incorrect password in C#
// Common Searches: how to test that a password-protected Excel file created with Aspose.Cells throws an exception when opened with a wrong password in C# | Aspose.Cells C# load encrypted workbook with wrong password handling | verify workbook password mismatch using LoadOptions in .NET | catch CellsException for invalid password when opening encrypted XLSX with Aspose.Cells | C# code to ensure Excel file encrypted with Aspose.Cells is not accessible with a different password
// Tags: Aspose.Cells password protection validation | LoadOptions wrong password exception | CellsException invalid password | C# temporary file deletion Aspose.Cells | XLSX encryption Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a temporary XLSX workbook, sets Settings.Password to "Password1", saves it, and then attempts to open the same file with LoadOptions.Password set to "Password2". Loading with the wrong password raises a CellsException, which is caught to confirm that the workbook cannot be opened with an incorrect password. The code also demonstrates proper cleanup of the temporary file.
class Program
{
    static void Main()
    {
        // Define a temporary file path for the encrypted workbook
        string filePath = Path.Combine(Path.GetTempPath(), "encryptedWorkbook.xlsx");

        try
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");

            // Encrypt the workbook with password "Password1"
            workbook.Settings.Password = "Password1";

            // Save the encrypted workbook
            workbook.Save(filePath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error creating or saving the workbook: " + ex.Message);
            return;
        }

        // Attempt to open the workbook with an incorrect password "Password2"
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
        {
            Password = "Password2"
        };

        try
        {
            // Ensure the file exists before attempting to load it
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"The file '{filePath}' was not found.");

            // This should throw a CellsException due to incorrect password
            Workbook loadedWorkbook = new Workbook(filePath, loadOptions);
            Console.WriteLine("Workbook opened with incorrect password (unexpected).");
        }
        catch (CellsException)
        {
            // Expected outcome: the workbook cannot be opened with the wrong password
            Console.WriteLine("Failed to open workbook with incorrect password as expected.");
        }
        catch (Exception ex)
        {
            // Any other exception is unexpected
            Console.WriteLine("Unexpected exception: " + ex.Message);
        }
        finally
        {
            // Clean up the temporary file
            try
            {
                if (File.Exists(filePath))
                    File.Delete(filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error deleting temporary file: " + ex.Message);
            }
        }
    }
}
