// Title: How to list all password‑protected worksheets in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# retrieve names of worksheets that require a password | list protected sheets in an .xlsx file with Aspose.Cells | C# code to identify password protected worksheets in a workbook | enumerate worksheets with protection enabled using Aspose.Cells .NET | how to check worksheet.IsProtected and password with Aspose.Cells
// Tags: Aspose.Cells worksheet protection detection | extract protected worksheet names .NET | worksheet.IsProtected password verification C# | generate protected sheet report .NET | Excel workbook protection status extraction

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The example loads an Excel workbook with Aspose.Cells, iterates through each Worksheet, checks the IsProtected flag and non‑empty Protection.Password, collects the names of all password‑protected sheets, and prints the list while handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        // Path to the workbook to analyze
        string workbookPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"File not found: {workbookPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Collect names of worksheets that are password protected
            List<string> protectedSheets = new List<string>();

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Worksheet is considered password protected if protection is enabled
                // and a password has been set (Password property is not empty)
                if (sheet.IsProtected && !string.IsNullOrEmpty(sheet.Protection.Password))
                {
                    protectedSheets.Add(sheet.Name);
                }
            }

            // Display the result
            Console.WriteLine("Password protected worksheets:");
            foreach (string name in protectedSheets)
            {
                Console.WriteLine(name);
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
