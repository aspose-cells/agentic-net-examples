// Title: Add purpose and business‑rule comments to Excel named ranges using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that loops through all workbook named ranges and sets a descriptive comment based on each range's name. | Show how to assign business‑rule text to specific named ranges (e.g., SalesData, RegionList, QuarterlyTargets) using the Name.Comment property in Aspose.Cells. | Provide an example that saves the modified workbook after adding comments to named ranges with Aspose.Cells for .NET.
// Common Searches: how to add comments to named ranges in an Excel file using Aspose.Cells C# | Aspose.Cells set Name.Comment property for specific named ranges | C# iterate workbook.Names and annotate each range with business rules | programmatically document Excel named ranges with purpose text using Aspose.Cells | save workbook after updating named range comments with Aspose.Cells .NET
// Tags: Aspose.Cells set named range comment | C# annotate Excel named ranges | Aspose.Cells add business rule notes to named ranges | Excel named range documentation with Aspose.Cells | update Name.Comment property Aspose.Cells .NET

using Aspose.Cells;
using System;

// The program loads an existing workbook, iterates through all defined names, assigns purpose‑specific comments (including business rules) to known ranges such as SalesData, RegionList, and QuarterlyTargets, applies a generic comment to others, and saves the updated file.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all defined names (named ranges) in the workbook
        foreach (Name namedRange in workbook.Worksheets.Names)
        {
            // Add a descriptive comment based on the name of the range
            switch (namedRange.Text)
            {
                case "SalesData":
                    // This range contains raw sales figures for the current fiscal year.
                    // Business rule: Values must be non‑negative and formatted as currency.
                    namedRange.Comment = "Contains raw sales figures for the current fiscal year. Values must be non-negative and formatted as currency.";
                    break;

                case "RegionList":
                    // List of valid sales regions.
                    // Business rule: Used for data validation in drop‑down lists.
                    namedRange.Comment = "List of valid sales regions. Used for data validation in drop-down lists.";
                    break;

                case "QuarterlyTargets":
                    // Target values for each quarter.
                    // Business rule: Targets cannot be lower than last year's corresponding quarter.
                    namedRange.Comment = "Target values for each quarter. Targets cannot be lower than last year's corresponding quarter.";
                    break;

                default:
                    // Generic comment for any other named range.
                    namedRange.Comment = "Named range for internal calculations.";
                    break;
            }
        }

        // Save the workbook with the added comments
        workbook.Save("output.xlsx");
    }
}
