// Title: Find hidden and protected worksheets in an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells that takes a file path and returns a list of worksheet names that are both hidden and protected. | Create a reusable method in C# that scans an Excel workbook with Aspose.Cells and prints each hidden protected sheet name. | Extend the sample to also output the protection status (e.g., password set) for each hidden worksheet using Aspose.Cells.
// Common Searches: asp.net how to list hidden protected worksheets using Aspose.Cells | c# Aspose.Cells retrieve names of hidden sheets that are password protected | detect hidden and protected worksheets in .xlsx with Aspose.Cells library | check worksheet visibility and protection status programmatically Aspose.Cells C#
// Tags: Aspose.Cells hidden worksheet detection | C# list protected Excel sheets | retrieve hidden protected worksheet names Aspose | Excel workbook worksheet visibility check Aspose.Cells | Aspose.Cells worksheet protection status .NET

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// Loads an Excel file with Aspose.Cells, iterates through all worksheets, checks IsVisible and IsProtected flags, collects names of sheets that are both hidden and protected, and outputs the list or a message when none are found.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string filePath = "input.xlsx";

            // Ensure the file exists before attempting to load it
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // List to hold names of worksheets that are both hidden and protected
            List<string> hiddenProtectedSheets = new List<string>();

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // A worksheet is hidden when its IsVisible property is false
                bool isHidden = !sheet.IsVisible;

                // Determine if the worksheet is protected.
                // In newer Aspose.Cells versions the Worksheet class exposes IsProtected directly.
                bool isProtected = sheet.IsProtected;

                // Add the sheet name to the list if both conditions are met
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
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
