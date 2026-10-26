// Title: Get defined names from an Excel workbook in C# with Aspose.Cells using a LoadFilter to skip worksheet data
// Common Searches: Aspose.Cells C# load filter to list only defined names | How to retrieve named ranges without loading worksheets in Aspose.Cells | C# example for extracting Excel named ranges using LoadFilter | Retrieve workbook names collection while avoiding sheet data with Aspose.Cells
// Tags: Aspose.Cells load filter defined names | C# retrieve named ranges Aspose.Cells | Excel named ranges without loading worksheets | Workbook.Names collection Aspose.Cells | load workbook without sheet data C#

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an Excel file with Aspose.Cells, accesses the Worksheets.Names collection to enumerate all defined names, and prints each name and its reference. By applying a LoadFilter you can prevent worksheet data from being loaded, making the operation faster and memory‑efficient.
class Program
{
    static void Main()
    {
        // Path to the Excel file
        string filePath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook without a custom load filter (optional)
            Workbook workbook = new Workbook(filePath);

            // Retrieve the collection of defined names
            NameCollection definedNames = workbook.Worksheets.Names;

            // Output each defined name and its reference
            foreach (Name name in definedNames)
            {
                Console.WriteLine($"Name: {name.Text}, RefersTo: {name.RefersTo}");
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
