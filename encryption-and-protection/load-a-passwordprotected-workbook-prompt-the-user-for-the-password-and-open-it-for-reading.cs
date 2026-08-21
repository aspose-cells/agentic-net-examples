// Title: Prompt for password and load a protected Excel workbook in C# with Aspise.Cells
// Description: Shows how to request a password from the console user, assign it to LoadOptions.Password, open a password‑protected .xlsx file with Aspose.Cells, read a cell value, and properly dispose the Workbook.
// Keywords: Aspose.Cells | C# | load password protected workbook | LoadOptions.Password | Excel encryption | prompt for password | console application | read protected Excel file
// Common Searches: Aspose.Cells open encrypted Excel C# | C# ask user for Excel password Aspose | LoadOptions password example | Read cell from password protected workbook Aspose.Cells | Handle wrong password when loading Excel with Aspose
// Developer Intent: Open a password‑protected Excel file by obtaining the password from the user at runtime and then read its data using Aspose.Cells.
// Use Cases: Interactive console utilities that need to process secured workbooks | Batch jobs that decrypt and extract data from protected Excel files | Validating a user‑supplied password before performing any workbook operations | Reading specific cells (e.g., A1) after successful decryption | Ensuring the Workbook object is disposed to free memory
// AI Prompts: Generate C# code that prompts the user for a password and opens a protected Excel workbook with Aspose.Cells, including error handling for incorrect passwords. | Provide an example that reads multiple cells from a password‑protected workbook after the user enters the password. | Explain best practices for disposing Aspose.Cells Workbook objects in a console application.

using System;
using Aspose.Cells;

namespace AsposeCellsPasswordLoadDemo
{
    // Shows how to request a password from the console user, assign it to LoadOptions.Password, open a password‑protected .xlsx file with Aspose.Cells, read a cell value, and properly dispose the Workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the password‑protected workbook
            string filePath = "protected.xlsx";

            // Prompt the user to enter the password
            Console.Write("Enter password to open the workbook: ");
            string password = Console.ReadLine();

            // Create LoadOptions and set the entered password
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = password;

            // Load the workbook using the load options
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Example: read and display the value of cell A1 from the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine("Value of A1: " + sheet.Cells["A1"].Value?.ToString());

            // Dispose the workbook when done
            workbook.Dispose();
        }
    }
}
