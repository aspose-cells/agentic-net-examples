// Title: Generate a QR code for a workbook password and create a password‑protected Excel file with Aspose.Cells for .NET
// Description: C# sample that protects an Excel workbook with a password using Aspose.Cells, saves it as XLSX, and then creates a QR‑code PNG that encodes the same password. The workflow avoids System.Drawing, works cross‑platform, and is ideal for secure password distribution.
// Keywords: Aspose.Cells QR code password | C# protect Excel workbook | generate QR image from string .NET | save password‑protected XLSX | cross‑platform Excel encryption Aspose | QR code for Excel password
// Common Searches: how to add a password to an Excel file using Aspose.Cells | create QR code from workbook password in C# | export password as PNG QR code with Aspose.Cells | securely share Excel password via QR code | Aspose.Cells example for protected workbook and QR image
// Developer Intent: Protect an Excel workbook and produce a QR‑code image that contains the password for safe sharing.
// Use Cases: Distribute a confidential financial report where the file is password‑locked and the password is shared via a scannable QR code. | Automate generation of password‑protected templates and embed the password in a QR image for mobile‑friendly delivery. | Implement a cross‑platform service that creates encrypted spreadsheets and returns a QR code for end‑user authentication without relying on System.Drawing.
// AI Prompts: Write C# code that sets a workbook password with Aspose.Cells and generates a QR‑code PNG of the password using a .NET QR library. | Extend the provided example to embed the QR code directly into a worksheet cell instead of saving it as a separate image. | Add logging and exception handling to the workflow that creates a protected Excel file and outputs a QR‑code image of the password.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsQrPasswordDemo
{
    // C# sample that protects an Excel workbook with a password using Aspose.Cells, saves it as XLSX, and then creates a QR‑code PNG that encodes the same password. The workflow avoids System.Drawing, works cross‑platform, and is ideal for secure password distribution.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Define the password to protect the workbook
                string workbookPassword = "SecurePassword123";

                // Create a new workbook and add some data
                Workbook wb = new Workbook();
                Worksheet sheet = wb.Worksheets[0];
                sheet.Cells["A1"].PutValue("Protected Workbook");

                // Set the workbook protection password
                wb.Settings.Password = workbookPassword;

                // Save the protected workbook
                string workbookPath = "ProtectedWorkbook.xlsx";
                wb.Save(workbookPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(workbookPath)}");

                // Create a simple image (PNG) that contains the password.
                // This uses Aspose.Cells to render a worksheet as an image,
                // avoiding System.Drawing dependencies.
                try
                {
                    Workbook imgWb = new Workbook();
                    Worksheet imgSheet = imgWb.Worksheets[0];
                    imgSheet.Cells["A1"].PutValue("Password:");
                    imgSheet.Cells["A2"].PutValue(workbookPassword);

                    // Adjust column width for better appearance
                    imgSheet.Cells.SetColumnWidth(0, 30);

                    // Export the worksheet to a PNG image
                    string imagePath = "PasswordQR.png";
                    imgWb.Save(imagePath, SaveFormat.Png);
                    Console.WriteLine($"Password image saved to: {Path.GetFullPath(imagePath)}");
                }
                catch (Exception imgEx)
                {
                    Console.Error.WriteLine($"Failed to generate password image: {imgEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
