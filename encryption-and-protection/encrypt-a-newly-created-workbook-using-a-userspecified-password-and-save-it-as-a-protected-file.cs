// Title: Encrypt a newly created Excel workbook with a custom password using Aspose.Cells for .NET and save as .xlsx
// AI Prompts: Generate a C# program that creates a Workbook, assigns a user‑provided string to workbook.Settings.Password, and saves the file as an encrypted .xlsx with Aspose.Cells. | Update the example to read the password from console input, apply it to workbook.Settings.Password, and then write the protected workbook to disk. | Adapt the code to encrypt the workbook as a binary .xlsb file while preserving password protection using Aspose.Cells.
// Common Searches: Aspose.Cells C# how to set a password for opening an Excel file | Encrypt newly created workbook with user defined password using Aspose.Cells .NET | Programmatically save a .xlsx file that requires a password to open using Aspose.Cells | C# example for workbook.Settings.Password property in Aspose.Cells
// Tags: Aspose.Cells workbook password protection | C# set workbook.Settings.Password | save password protected .xlsx Aspose.Cells | create encrypted Excel file .NET | protect Excel workbook with user password

using System;
using Aspose.Cells;

// // This example creates a new Workbook, adds sample data, assigns a user‑specified password via workbook.Settings.Password to encrypt the file, and saves the protected workbook as ProtectedWorkbook.xlsx in Xlsx format.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // (Optional) Add some sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample");
        sheet.Cells["B1"].PutValue("Data");

        // Set the password that will be required to open the file
        // This encrypts the workbook using the specified password
        workbook.Settings.Password = "UserSpecifiedPassword";

        // Save the encrypted workbook to disk
        workbook.Save("ProtectedWorkbook.xlsx", SaveFormat.Xlsx);
    }
}
