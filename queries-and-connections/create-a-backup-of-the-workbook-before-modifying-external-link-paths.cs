// Title: How to create a backup of an Excel workbook before modifying external link paths with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells save backup of workbook before changing external links | Aspose.Cells .NET create copy of Excel file prior to updating external link paths | preserve original Excel file when editing external links using Aspose.Cells | backup original.xlsx then modify external link paths with Aspose.Cells in C# | Aspose.Cells example for saving workbook backup and then updating external links
// Tags: Aspose.Cells create workbook backup C# | modify external link paths Aspose.Cells | Aspose.Cells SaveFormat.Xlsx usage | preserve original Excel file Aspose.Cells | copy workbook before external link update Aspose.Cells

using Aspose.Cells;
using System;

// The sample loads an existing Excel workbook, saves a backup copy named "original_backup.xlsx" using Aspose.Cells, provides a placeholder for iterating and updating external link paths, and finally saves the modified workbook as "original_modified.xlsx".
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("original.xlsx");

        // Create a backup copy before making any changes
        workbook.Save("original_backup.xlsx", SaveFormat.Xlsx);

        // -------------------------------------------------
        // Modify external link paths here (example placeholder)
        // foreach (ExternalLink link in workbook.Worksheets.ExternalLinks)
        // {
        //     // Update link.Path or other properties as needed
        // }
        // -------------------------------------------------

        // Save the workbook after modifications
        workbook.Save("original_modified.xlsx", SaveFormat.Xlsx);
    }
}
