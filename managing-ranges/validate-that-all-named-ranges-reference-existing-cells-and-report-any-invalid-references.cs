// Title: Validate named ranges and identify invalid references in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# code to verify that all named ranges in an Excel file exist with Aspose.Cells | how to list invalid named ranges in a workbook using Aspose.Cells .NET | detect broken named range references in Excel using Aspose.Cells API | Aspose.Cells GetRangeByName returns null for missing range example | validate defined names integrity in an Excel workbook with Aspose.Cells for .NET
// Tags: named range validation Aspose.Cells | GetRangeByName null check | detect broken defined names C# | Excel workbook named range integrity | report invalid named ranges Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// Alias to avoid conflict with System.Range introduced in C# 8.0
using AsposeRange = Aspose.Cells.Range;

// The example loads an Excel workbook, iterates through its NameCollection, uses GetRangeByName to determine if each named range points to an existing cell range, and outputs messages for any names that reference non‑existent cells or cause errors, otherwise confirming all named ranges are valid.
class NamedRangeValidator
{
    static void Main(string[] args)
    {
        // Path to the Excel file to validate
        string filePath = "input.xlsx";

        // Ensure the file exists before attempting to load
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(filePath);

            // Retrieve the collection of all defined names (named ranges) in the workbook
            NameCollection names = workbook.Worksheets.Names;

            bool hasInvalid = false;

            // Iterate through each named range and check its validity
            foreach (Name name in names)
            {
                try
                {
                    // Get the range referenced by the name; null indicates an invalid reference
                    AsposeRange range = workbook.Worksheets.GetRangeByName(name.Text);
                    if (range == null)
                    {
                        hasInvalid = true;
                        Console.WriteLine($"Invalid named range: '{name.Text}' refers to '{name.RefersTo}' which does not exist.");
                    }
                }
                catch (Exception ex)
                {
                    // Handle errors specific to a single named range
                    hasInvalid = true;
                    Console.WriteLine($"Error checking named range '{name.Text}': {ex.Message}");
                }
            }

            if (!hasInvalid)
            {
                Console.WriteLine("All named ranges are valid.");
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"Error processing workbook: {ex.Message}");
        }
    }
}
