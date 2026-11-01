// Title: Find hidden and password‑protected worksheets in an Excel file using Aspose.Cells for .NET
// Common Searches: C# code to list hidden worksheets that are also protected with Aspose.Cells | how to check if an Excel sheet is hidden and protected using Aspose.Cells .NET | retrieve names of hidden protected worksheets from a workbook in C# | Aspose.Cells detect hidden and protected worksheets in .xlsx file | enumerate hidden protected sheets in Excel using Aspose.Cells API
// Tags: Aspose.Cells hidden worksheet detection | Aspose.Cells protected sheet enumeration | C# worksheet IsVisible IsProtected check | Excel hidden protected sheet listing .NET | Aspose.Cells workbook worksheet protection status

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, iterates through each worksheet, checks the IsVisible and IsProtected properties, collects the names of sheets that are both hidden and protected, and prints the list or a message if none are found.
class Program
{
    static void Main()
    {
        // Path to the input workbook
        string filePath = "input.xlsx";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // List to hold names of hidden and protected worksheets
            List<string> hiddenProtectedSheets = new List<string>();

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine if the worksheet is hidden
                bool isHidden = !sheet.IsVisible;

                // Determine if the worksheet is protected
                bool isProtected = sheet.IsProtected;

                // If both conditions are true, add the sheet name to the list
                if (isHidden && isProtected)
                {
                    hiddenProtectedSheets.Add(sheet.Name);
                }
            }

            // Output the results
            if (hiddenProtectedSheets.Count > 0)
            {
                Console.WriteLine("Hidden and protected worksheets found:");
                foreach (string name in hiddenProtectedSheets)
                {
                    Console.WriteLine("- " + name);
                }
            }
            else
            {
                Console.WriteLine("No hidden protected worksheets found.");
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
