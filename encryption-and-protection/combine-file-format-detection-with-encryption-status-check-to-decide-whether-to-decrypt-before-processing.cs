// Title: Detect Excel file format and conditionally decrypt password‑protected workbooks with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel file using Aspose.Cells LoadOptions with LoadFormat.Auto, tries to open it without a password, and if a CellsException occurs, reloads the workbook with a supplied password. | Create a reusable C# method that takes a file path and password, determines whether the workbook is encrypted, and returns an opened Workbook instance ready for further processing. | Show how to read the value of cell A1 after handling possible encryption and then save the workbook to a new file using Aspose.Cells.
// Common Searches: aspocells c# load excel file with automatic format detection and password fallback | how to open encrypted .xlsx using Aspose.Cells LoadOptions in .NET | detect file format and decrypt Excel workbook programmatically with Aspose.Cells | c# example for handling password protected Excel files with Aspose.Cells
// Tags: auto format detection LoadOptions Aspose.Cells | conditional workbook decryption C# | open password protected Excel with Aspose.Cells | read cell A1 after workbook decryption .NET | save processed workbook Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The program uses Aspose.Cells to automatically detect the Excel file format, attempts to load the workbook without a password, and if a CellsException indicates encryption, reloads it with the provided password before reading cell A1 and saving the result as a new file.
class Program
{
    static void Main()
    {
        // Input file path (provide full name with extension)
        string inputPath = "inputFile.xlsx";

        // Password for encrypted workbooks (if needed)
        string password = "myPassword";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        // Load options with automatic format detection
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
        Workbook workbook = null;

        try
        {
            // Attempt to load the workbook without a password first
            workbook = new Workbook(inputPath, loadOptions);
        }
        catch (CellsException)
        {
            // Workbook may be encrypted; reload using the supplied password
            loadOptions.Password = password;
            try
            {
                workbook = new Workbook(inputPath, loadOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load encrypted workbook: {ex.Message}");
                return;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        try
        {
            // Example processing: read value from first worksheet cell A1
            Worksheet sheet = workbook.Worksheets[0];
            string cellValue = sheet.Cells["A1"].StringValue;
            Console.WriteLine($"Value in A1: {cellValue}");

            // (Additional processing can be added here)

            // Save the processed workbook
            workbook.Save("output.xlsx");
            Console.WriteLine("Workbook saved as output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during processing or saving: {ex.Message}");
        }
    }
}
