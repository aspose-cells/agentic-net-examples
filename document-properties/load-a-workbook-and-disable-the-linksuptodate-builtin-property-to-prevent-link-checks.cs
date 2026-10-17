// Title: Disable the LinksUpToDate built‑in document property in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set LinksUpToDate property to false | prevent Excel external link validation using Aspose.Cells | how to turn off link update checks in a workbook with Aspose.Cells | disable built‑in LinksUpToDate property when saving Excel file in .NET
// Tags: turn off workbook link status flag Aspose.Cells | modify built‑in document properties C# | suppress Excel external link validation Aspose.Cells | C# set workbook property false Aspose.Cells | Excel link check suppression using Aspose.Cells

using Aspose.Cells;
using System;

// Loads an existing workbook, sets the LinksUpToDate built‑in document property to false to stop link validation, and saves the workbook.
class Program
{
    static void Main()
    {
        // Load the existing workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Disable the LinksUpToDate built‑in property to prevent link checks
        workbook.BuiltInDocumentProperties.LinksUpToDate = false;

        // Save the workbook (optional, depending on whether you want to persist the change)
        workbook.Save("output.xlsx");
    }
}
