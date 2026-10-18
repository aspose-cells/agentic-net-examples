// Title: Encrypt a new workbook with a custom password and save it as a protected XLSX using Aspose.Cells for .NET
// Common Searches: how to set an open password for a new Excel workbook using Aspose.Cells C# | Aspose.Cells C# encrypt workbook and protect its structure before saving | save workbook with password protection and encryption in Aspose.Cells .NET | C# example for workbook.Settings.Password and workbook.Protect in Aspose.Cells | prevent modifications to an Excel file with password using Aspose.Cells
// Tags: Aspose.Cells workbook password encryption | C# set workbook open password | protect workbook structure Aspose.Cells | save encrypted XLSX with Aspose.Cells | Aspose.Cells workbook protection API

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new Workbook, writes sample data to cells A1 and B1, assigns an open password via workbook.Settings.Password, applies full protection (structure and windows) with workbook.Protect, and saves the encrypted and protected file as 'ProtectedWorkbook.xlsx' in XLSX format using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample");
            sheet.Cells["B1"].PutValue("Data");

            // Set the password required to open the file
            workbook.Settings.Password = "UserSpecifiedPassword";

            // Protect the workbook from modifications (structure and windows)
            workbook.Protect(ProtectionType.All, "UserSpecifiedPassword");

            // Define output file path
            string outputPath = "ProtectedWorkbook.xlsx";

            // Save the encrypted and protected workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
