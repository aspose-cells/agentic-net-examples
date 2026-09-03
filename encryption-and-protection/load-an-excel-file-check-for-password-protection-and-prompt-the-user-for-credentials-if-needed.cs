// Title: C# Aspose.Cells: Open an .xlsx file and request a password when the workbook is password‑protected
// AI Prompts: Write C# code that uses Aspose.Cells to open an .xlsx file, catches the password‑required exception, prompts the user for the password via the console, and reloads the workbook with LoadOptions.Password. | Show how to detect a protected Excel workbook with Aspose.Cells, ask the user for credentials, and load the file using the supplied password in C#.
// Common Searches: aspocells c# load password protected xlsx and ask user for password | how to handle CellsException for encrypted workbook in Aspose.Cells | load excel file with user supplied password using Aspose.Cells C# | catch password error when opening .xlsx with Aspose.Cells and prompt for input
// Tags: Aspose.Cells load password protected workbook C# | LoadOptions.Password Aspose.Cells | catch CellsException password protection | prompt user for Excel file password C# | open encrypted .xlsx with Aspose.Cells

using System;
using Aspose.Cells;

// The program attempts to open 'input.xlsx' with Aspose.Cells; if a CellsException indicating a password is caught, it reads a password from the console, reloads the workbook using LoadOptions.Password, and confirms successful loading.
class Program
{
    static void Main()
    {
        // Path to the Excel file
        string filePath = "input.xlsx";

        Workbook workbook = null;

        try
        {
            // Attempt to load the workbook without a password
            workbook = new Workbook(filePath);
        }
        catch (CellsException ex) when (ex.Message.Contains("Password"))
        {
            // The file is password protected; ask the user for the password
            Console.Write("Enter password for the workbook: ");
            string password = Console.ReadLine();

            // Load the workbook using the supplied password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };
            workbook = new Workbook(filePath, loadOptions);
        }

        // At this point the workbook is loaded (either without password or with the provided one)
        Console.WriteLine("Workbook loaded successfully.");
        // Further processing can be done here...
    }
}
