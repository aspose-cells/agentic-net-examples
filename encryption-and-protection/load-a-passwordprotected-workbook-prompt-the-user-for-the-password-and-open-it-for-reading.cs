// Title: Open a Password‑Protected Excel Workbook in C# with Aspose.Cells – Prompt for Password
// Description: A console program that asks the user for a password, applies it via LoadOptions, and opens a protected .xlsx file with Aspose.Cells. After loading, it reads cell A1 from the first worksheet and prints the value.
// Keywords: Aspose.Cells password protected workbook | C# load encrypted Excel file | LoadOptions password | read protected .xlsx | prompt user for Excel password
// Common Searches: how to open a password protected Excel file using Aspose.Cells C# | Aspose.Cells LoadOptions password example | read cell from encrypted workbook .NET | prompt for workbook password Aspose.Cells
// Developer Intent: Open a secured Excel file after obtaining the password from the user and access its data.
// Use Cases: Securely load a protected workbook by requesting the password at runtime. | Validate the entered password and handle incorrect entries gracefully. | Extract specific cells or iterate through worksheets in an encrypted file.
// AI Prompts: Generate C# code that uses Aspose.Cells to open a password‑protected workbook after asking the user for the password, including exception handling for wrong passwords. | Create a reusable method that receives a file path, prompts for a password, loads the workbook with LoadOptions, and returns the value of a specified cell. | Show how to list all worksheet names from an encrypted .xlsx using Aspose.Cells in C#.

using System;
using Aspose.Cells;

// A console program that asks the user for a password, applies it via LoadOptions, and opens a protected .xlsx file with Aspose.Cells. After loading, it reads cell A1 from the first worksheet and prints the value.
class Program
{
    static void Main()
    {
        // Path to the password‑protected workbook
        string filePath = "protected.xlsx";

        // Prompt the user to enter the password
        Console.Write("Enter password to open the workbook: ");
        string password = Console.ReadLine();

        // Create LoadOptions and set the password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = password;

        // Load the workbook using the provided password
        Workbook workbook = new Workbook(filePath, loadOptions);

        // Example: read a value from the first worksheet
        string cellValue = workbook.Worksheets[0].Cells["A1"].StringValue;
        Console.WriteLine($"Value of A1: {cellValue}");
    }
}
