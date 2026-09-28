// Title: Remove the custom document property 'IsReviewed' from an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Create C# code using Aspose.Cells that opens a given .xlsx file, checks for a custom document property named 'IsReviewed', deletes it if it exists, and saves the workbook. | Write a reusable C# method that accepts a file path and a property name, removes that custom document property from the workbook using Aspose.Cells, and returns the updated file.
// Common Searches: asp.net remove custom document property IsReviewed from Excel using aspose.cells | c# delete specific custom property from .xlsx workbook programmatically | how to clean obsolete custom properties in an Excel file with Aspose.Cells | remove metadata property from Excel workbook C# Aspose.Cells example
// Tags: remove custom document property Aspose.Cells | delete Excel custom property C# | Aspose.Cells workbook metadata cleanup | C# remove specific custom property .xlsx | custom property management Aspose.Cells

using System;
using Aspose.Cells;

// The example loads an Excel file with Aspose.Cells, checks whether the custom document property 'IsReviewed' exists, removes it if present, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Check if the custom property "IsReviewed" exists and remove it
        if (workbook.CustomDocumentProperties.Contains("IsReviewed"))
        {
            workbook.CustomDocumentProperties.Remove("IsReviewed");
        }

        // Save the workbook after removing the property
        workbook.Save("output.xlsx");
    }
}
