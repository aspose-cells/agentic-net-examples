// Title: Detect hidden and protected worksheets in an Excel file and list their names using Aspose.Cells for .NET (C#)
// Common Searches: C# code to find hidden and password‑protected worksheets with Aspose.Cells | list names of very hidden protected sheets in an .xlsx using Aspose.Cells | how to check worksheet IsVisible and IsProtected properties in Aspose.Cells | identify hidden protected worksheets in Excel workbook via Aspose.Cells .NET | retrieve hidden protected sheet names from workbook using Aspose.Cells API
// Tags: hidden protected worksheet detection Aspose.Cells | worksheet IsProtected property Aspose.Cells | worksheet visibility check Aspose.Cells | enumerate very hidden sheets .NET | list protected sheet names C# Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The example loads an Excel workbook with Aspose.Cells, iterates through each worksheet, checks both the visibility (hidden or very hidden) and protection status, collects the names of worksheets that satisfy both conditions, and outputs the list or a message if none are found.
class Program
{
    static void Main()
    {
        // Path to the input workbook
        string filePath = "input.xlsx";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook from the specified file
            workbook = new Workbook(filePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // List to hold names of hidden and protected worksheets
        List<string> hiddenProtectedSheets = new List<string>();

        // Examine each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Determine if the sheet is hidden (Hidden or VeryHidden)
            bool isHidden = !sheet.IsVisible;

            // Determine if the sheet is protected
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
}
