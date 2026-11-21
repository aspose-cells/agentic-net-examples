// Title: How to extract defined names (named ranges) from an Excel workbook using Aspose.Cells for .NET in C#
// Common Searches: C# Aspose.Cells read all defined names from an .xlsx workbook | list named ranges in Excel file using Aspose.Cells API | retrieve NameCollection from Workbook.Worksheets.Names in C# | extract name.Text and name.RefersTo values with Aspose.Cells | Aspose.Cells get named range references programmatically
// Tags: Aspose.Cells defined names extraction | C# Excel named ranges retrieval | Workbook.Worksheets.Names usage | extract name.RefersTo property | list Excel named ranges Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, accesses the NameCollection via Workbook.Worksheets.Names, iterates through each defined name, combines its Text and RefersTo properties into a readable string, stores the results in a list, and outputs each entry to the console.
class Program
{
    static void Main()
    {
        const string workbookPath = "input.xlsx";

        // Verify that the workbook file exists to avoid FileNotFoundException
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: The file \"{workbookPath}\" was not found.");
            return;
        }

        try
        {
            // Load the Excel workbook
            Workbook workbook = new Workbook(workbookPath);

            // Retrieve the collection of defined names (named ranges)
            NameCollection definedNames = workbook.Worksheets.Names;

            // Prepare a list to hold the extracted information
            List<string> nameInfoList = new List<string>();

            // Iterate through each defined name and store its name and reference
            foreach (Name name in definedNames)
            {
                // Example format: "MyRange - Sheet1!$A$1:$B$2"
                string entry = $"{name.Text} - {name.RefersTo}";
                nameInfoList.Add(entry);
            }

            // Output the extracted defined names
            foreach (string entry in nameInfoList)
            {
                Console.WriteLine(entry);
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
