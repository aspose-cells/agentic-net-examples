// Title: Prefix custom XML part names with a project identifier in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add project prefix to custom XML part names in an existing .xlsx file | rename custom XML parts in Excel workbook programmatically with Aspose.Cells .NET | apply naming convention to workbook custom XML parts using Aspose.Cells API | how to iterate and rename custom XML parts in a spreadsheet with Aspose.Cells C# | prefix custom XML part names with project code when saving workbook using Aspose.Cells
// Tags: Aspose.Cells custom XML part renaming | C# prefix custom XML part names | Excel workbook custom XML parts manipulation | project identifier naming convention Aspose.Cells | rename custom XML parts in .xlsx using .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing .xlsx workbook with Aspose.Cells for .NET, iterates through its CustomXmlParts collection, prepends a specified project identifier to each part’s Name property, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Input workbook path
            string inputPath = "input.xlsx";
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load workbook
            var workbook = new Workbook(inputPath);

            // Prefix for custom XML part names
            string projectId = "Proj123";

            // Rename each custom XML part with the project prefix
            for (int i = 0; i < workbook.CustomXmlParts.Count; i++)
            {
                dynamic xmlPart = workbook.CustomXmlParts[i];
                string originalName = xmlPart.Name as string;
                xmlPart.Name = $"{projectId}_{originalName}";
            }

            // Save modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
