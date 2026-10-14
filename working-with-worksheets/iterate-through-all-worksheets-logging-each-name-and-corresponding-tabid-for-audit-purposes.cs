// Title: Log each worksheet's name and TabId while iterating through all worksheets in an Aspose.Cells workbook (C#)
// AI Prompts: Write C# code using Aspose.Cells that enumerates every worksheet in a workbook and prints both the worksheet Name and its TabId to the console. | Create a C# method that accepts a Workbook object and returns a list of objects containing worksheet Name, TabId, and Visibility status using Aspose.Cells. | Generate a reusable utility in C# that builds a dictionary mapping worksheet names to their TabId values from an Aspose.Cells workbook.
// Common Searches: how to print worksheet TabId with Aspose.Cells in C# | C# Aspose.Cells enumerate worksheets and display their names and TabIds | retrieve worksheet identifiers for audit using Aspose.Cells .NET | list worksheet visibility and TabId using Aspose.Cells C#
// Tags: Aspose.Cells iterate worksheets C# | Aspose.Cells get worksheet TabId | audit Excel sheet identifiers .NET | log worksheet name and TabId C# | retrieve worksheet visibility Aspose.Cells

using System;
using Aspose.Cells;

// Loads an Excel file with Aspose.Cells, loops through each worksheet, and writes the worksheet's Name and TabId to the console for auditing purposes.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Log worksheet name and its TabId
            Console.WriteLine($"Worksheet Name: {sheet.Name}, TabId: {sheet.TabId}");
        }

        // Optionally save the workbook if any changes were made
        // workbook.Save("output.xlsx");
    }
}
