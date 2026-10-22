// Title: Add purpose and business‑rule comments to every named range in an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add comment to each named range in an existing workbook | How to programmatically document Excel named ranges using Aspose.Cells | Set purpose and business rule metadata for named ranges with Aspose.Cells .NET | Iterate through workbook.Worksheets.Names and assign comments in C# | Aspose.Cells example for annotating named ranges with custom text | Add descriptive comments to Excel named ranges via Aspose.Cells API
// Tags: Aspose.Cells add named range comment C# | C# annotate Excel named ranges with Aspose.Cells | Aspose.Cells set Name.Comment property | Excel named range metadata automation Aspose.Cells | Aspose.Cells workbook named range documentation | C# programmatic Excel named range annotation

using System;
using System.IO;
using Aspose.Cells;

// The C# sample loads an Excel workbook (creating a new file if missing), loops through all defined named ranges via workbook.Worksheets.Names, assigns a formatted comment that describes each range's purpose and associated business rules using the Name.Comment property, and saves the updated workbook with Aspose.Cells.
class Program
{
    static void Main()
    {
        try
        {
            // Load the workbook (creates a new one if the file does not exist)
            Workbook workbook = LoadWorkbook("input.xlsx");

            // Iterate through all named ranges and add descriptive comments
            foreach (Name namedRange in workbook.Worksheets.Names)
            {
                // Example comment: describe the purpose of the range and any business rules
                namedRange.Comment = $"Purpose: {namedRange.Text} is used for ...; Business Rules: ...";
            }

            // Save the workbook
            SaveWorkbook(workbook, "output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Loads a workbook from the specified path; creates a new workbook if the file is missing
    private static Workbook LoadWorkbook(string filePath)
    {
        if (File.Exists(filePath))
        {
            return new Workbook(filePath);
        }
        else
        {
            // Return an empty workbook as a fallback
            return new Workbook();
        }
    }

    // Saves the workbook to the specified path in XLSX format
    private static void SaveWorkbook(Workbook workbook, string filePath)
    {
        workbook.Save(filePath, SaveFormat.Xlsx);
    }
}
