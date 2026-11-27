// Title: C# example: verify workbook password protection flag after saving an encrypted Excel file with Aspose.Cells
// Common Searches: how to confirm that an Excel workbook is password protected after saving with Aspose.Cells in C# | Aspose.Cells C# load encrypted .xlsx and check Settings.Password flag | verify workbook protection status programmatically using Aspose.Cells LoadOptions | C# code to test if a saved Excel file has password protection enabled with Aspose.Cells | determine whether a workbook saved by Aspose.Cells is marked as password protected
// Tags: Aspose.Cells workbook password protection verification | C# load encrypted Excel with LoadOptions | check Settings.Password flag after save | validate encrypted workbook flag Aspose.Cells | programmatic detection of Excel file protection .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a Workbook, sets a password via wb.Settings.Password, saves it as an Xlsx file, then reloads the file using LoadOptions with the same password. It confirms the protection by verifying that loadedWb.Settings.Password is not empty, outputting a true flag indicating the workbook is password protected.
class VerifyPasswordProtection
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add some data
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Test");

            // Set password protection for the workbook
            string password = "Secret123";
            wb.Settings.Password = password; // marks workbook as password protected

            // Save the workbook to a file (the password is applied automatically)
            string filePath = "ProtectedWorkbook.xlsx";
            wb.Save(filePath, SaveFormat.Xlsx);

            // Ensure the file exists before attempting to load it
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"The file '{filePath}' was not found.");

            // Load the saved workbook using the password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };
            Workbook loadedWb = new Workbook(filePath, loadOptions);

            // Verify that the workbook was opened with a password
            bool isProtected = !string.IsNullOrEmpty(loadedWb.Settings.Password);

            Console.WriteLine($"Workbook password protection flag is set: {isProtected}");
            // Expected output: Workbook password protection flag is set: True
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
