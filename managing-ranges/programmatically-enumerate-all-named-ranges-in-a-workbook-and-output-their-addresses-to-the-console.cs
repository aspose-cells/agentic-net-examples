// Title: How to list all named ranges in an Excel workbook and print their addresses using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells list all defined names in an Excel file | retrieve named range addresses from a workbook with Aspose.Cells .NET | enumerate workbook named ranges and display RefersTo values in console | Aspose.Cells .NET get named range address programmatically | list named ranges in password protected Excel using Aspose.Cells C#
// Tags: Aspose.Cells list defined names | C# enumerate named ranges | Aspose.Cells get RefersTo address | Excel workbook named range extraction .NET | handle password protected workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook (including optional password handling), iterates through Workbook.Worksheets.Names, and writes each named range's name and its RefersTo address (e.g., =Sheet1!$A$1:$B$5) to the console, with robust error handling for file access and processing.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: File not found – \"{inputPath}\"");
            return;
        }

        Workbook workbook = null;
        LoadOptions loadOptions = new LoadOptions();

        try
        {
            // Attempt to load the workbook without a password
            workbook = new Workbook(inputPath, loadOptions);
        }
        catch (CellsException ex) when (ex.Message.Contains("password", StringComparison.OrdinalIgnoreCase))
        {
            // Workbook is password‑protected; request password from the user
            Console.Write("The workbook is password protected. Enter password: ");
            string password = Console.ReadLine() ?? string.Empty;

            loadOptions.Password = password;

            try
            {
                workbook = new Workbook(inputPath, loadOptions);
            }
            catch (Exception innerEx)
            {
                Console.WriteLine($"Failed to open workbook with provided password: {innerEx.Message}");
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
            // Iterate through all defined names (named ranges) in the workbook
            foreach (Name namedRange in workbook.Worksheets.Names)
            {
                // The RefersTo property contains the address the name points to (e.g., =Sheet1!$A$1:$B$5)
                string address = namedRange.RefersTo;
                Console.WriteLine($"{namedRange.Text}: {address}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error while processing named ranges: {ex.Message}");
        }
    }
}
