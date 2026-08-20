// Title: C# – Detect and Open Password‑Protected Excel Files with Aspose.Cells
// Description: A concise guide showing how to use Aspose.Cells to determine whether an Excel workbook is encrypted, request a password from the console when required, and load the file with LoadOptions. The sample also reads cell A1 after opening.
// Keywords: Aspose.Cells C# | detect encrypted workbook | password protected Excel file | FileFormatUtil DetectFileFormat | LoadOptions password | console password prompt | open Excel workbook | read cell value | Excel encryption detection
// Common Searches: Aspose.Cells check if Excel file is password protected | C# load encrypted workbook with user‑entered password | How to prompt for Excel password using Aspose.Cells | Detect encryption before opening Excel in .NET | Read cell after opening password‑protected Excel
// Developer Intent: Find out whether an Excel file is encrypted and load it, asking the user for a password only when necessary.
// Use Cases: Interactive console tools that need to open unknown Excel files safely. | Batch processing pipelines that skip or decrypt files based on encryption status. | Securely prompting end‑users for passwords before accessing protected worksheets.
// AI Prompts: Generate C# code that uses Aspose.Cells to detect an encrypted Excel workbook, asks the console user for the password if needed, and returns a Workbook object. | Create a reusable function in .NET that accepts a file path, checks for encryption with FileFormatUtil, prompts for a password when required, and loads the workbook with LoadOptions.

using System;
using Aspose.Cells;

namespace AsposeCellsPasswordDemo
{
    // A concise guide showing how to use Aspose.Cells to determine whether an Excel workbook is encrypted, request a password from the console when required, and load the file with LoadOptions. The sample also reads cell A1 after opening.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file
            string filePath = "sample.xlsx";

            // Detect file format and check if the workbook is encrypted
            FileFormatInfo fileInfo = FileFormatUtil.DetectFileFormat(filePath);
            Workbook workbook;

            if (fileInfo.IsEncrypted)
            {
                // Prompt the user for the password
                Console.Write("The workbook is password protected. Enter password: ");
                string password = Console.ReadLine();

                // Load the workbook using LoadOptions with the supplied password
                LoadOptions loadOptions = new LoadOptions();
                loadOptions.Password = password;
                workbook = new Workbook(filePath, loadOptions);
            }
            else
            {
                // Load the workbook normally (no password required)
                workbook = new Workbook(filePath);
            }

            // Example usage: output the value of cell A1 from the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine("Value of A1: " + sheet.Cells["A1"].Value?.ToString());

            // Keep console window open
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }
}
