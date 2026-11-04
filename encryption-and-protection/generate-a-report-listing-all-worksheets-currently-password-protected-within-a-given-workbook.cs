// Title: List password‑protected worksheets in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# list worksheets with a protection password | Get names of password‑protected sheets from an .xlsx file using Aspose.Cells | How to detect protected worksheets in a workbook with Aspose.Cells .NET | C# enumerate worksheets where IsProtected is true and password is set | Retrieve protected worksheet names from Excel using Aspose.Cells API
// Tags: enumerate protected worksheets Aspose.Cells | worksheet password detection .NET | list locked sheets Excel Aspose.Cells | retrieve protected sheet names C# | Aspose.Cells workbook protection inspection

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an .xlsx file with Aspose.Cells, iterates through all worksheets, checks each sheet’s IsProtected flag and password property, collects the names of sheets that are password‑protected, and prints a report of those worksheet names.
class Program
{
    static void Main()
    {
        // Path to the workbook to be examined
        string workbookPath = "input.xlsx";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: The file \"{workbookPath}\" was not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(workbookPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // List to hold names of password‑protected worksheets
        List<string> protectedSheets = new List<string>();

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Check if the worksheet is protected and a password has been set
            if (sheet.IsProtected && !string.IsNullOrEmpty(sheet.Protection.Password))
            {
                protectedSheets.Add(sheet.Name);
            }
        }

        // Output the result
        Console.WriteLine("Password‑protected worksheets:");
        if (protectedSheets.Count == 0)
        {
            Console.WriteLine("None");
        }
        else
        {
            foreach (string name in protectedSheets)
            {
                Console.WriteLine("- " + name);
            }
        }
    }
}
