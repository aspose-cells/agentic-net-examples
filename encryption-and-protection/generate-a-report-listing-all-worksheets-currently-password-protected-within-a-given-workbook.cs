// Title: Generate a report of all password‑protected worksheets in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, iterates through each worksheet, checks the Protection.Password property, and collects the names of protected sheets. | Add logic to output the list of password‑protected worksheet names to the console and then write them to a text file named ProtectedWorksheetsReport.txt. | Include error handling for missing workbook files and for exceptions that may occur while accessing worksheet protection settings.
// Common Searches: how to list password protected worksheets in an Excel file using Aspose.Cells C# | Aspose.Cells enumerate protected sheets and export names to a file | C# code to detect worksheets with a set protection password in .xlsx | save list of protected Excel worksheets to text file with Aspose.Cells
// Tags: enumerate worksheet protection Aspose.Cells | list password protected Excel sheets C# | export protected worksheet names to text file | handle missing workbook file Aspose.Cells | check Worksheet.Protection.Password property

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads a specified .xlsx workbook with Aspose.Cells, loops through all worksheets, examines each sheet's Protection.Password property, gathers the names of sheets that have a password, prints the list to the console, and writes the names to ProtectedWorksheetsReport.txt while handling file‑not‑found and protection‑checking exceptions.
class Program
{
    static void Main()
    {
        // Path to the workbook to be examined
        string workbookPath = "input.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: The file \"{workbookPath}\" was not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook using Aspose.Cells
            workbook = new Workbook(workbookPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Collection to store names of worksheets that are password protected
        List<string> protectedWorksheetNames = new List<string>();

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // If a password is set, the worksheet is considered protected
            try
            {
                if (!string.IsNullOrEmpty(sheet.Protection.Password))
                {
                    protectedWorksheetNames.Add(sheet.Name);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error checking protection for sheet \"{sheet.Name}\": {ex.Message}");
            }
        }

        // Generate the report: list worksheet names that are password protected
        Console.WriteLine("Password protected worksheets:");
        foreach (string name in protectedWorksheetNames)
        {
            Console.WriteLine("- " + name);
        }

        // Optionally save the report to a text file
        try
        {
            File.WriteAllLines("ProtectedWorksheetsReport.txt", protectedWorksheetNames);
            Console.WriteLine("Report saved to ProtectedWorksheetsReport.txt");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to write report file: {ex.Message}");
        }
    }
}
