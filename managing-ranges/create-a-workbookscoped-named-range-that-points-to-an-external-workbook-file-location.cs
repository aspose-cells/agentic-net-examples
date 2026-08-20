// Title: Define a Workbook‑Scoped Named Range that References an External Workbook with AspNet Aspose.Cells
// Description: Demonstrates how to create a primary workbook, add an external link to another .xlsx file (using a relative or absolute path), set the link’s data source, create a workbook‑scoped named range whose RefersTo formula points to a range in the linked workbook, and save the result.
// Keywords: Aspose.Cells external named range | workbook scoped name .NET | add external link Aspose.Cells | reference external workbook range | C# Aspose.Cells external file | named range RefersTo formula
// Common Searches: Aspose.Cells create workbook scoped named range to external file | C# add external link and named range in Aspose.Cells | How to set RefersTo for an external workbook range using Aspose.Cells | Define external named range in .NET Excel library
// Developer Intent: Create a workbook‑scoped named range that points to a range in an external Excel workbook using Aspose.Cells for .NET.
// Use Cases: Consolidate data by linking a summary workbook to a master data workbook via an external named range. | Build a reporting template that pulls specific cells from a shared source workbook without duplicating the source file. | Maintain a single source of truth across multiple workbooks by reusing external named ranges for common data blocks.
// AI Prompts: Generate C# code with Aspose.Cells to add an external link and define a workbook‑scoped named range that references a range in the linked workbook. | Show how to modify the RefersTo formula to target a different sheet or cell block in the external workbook. | Explain how to update the file path of an existing external named range in an Aspose.Cells workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsExternalNamedRange
{
    // Demonstrates how to create a primary workbook, add an external link to another .xlsx file (using a relative or absolute path), set the link’s data source, create a workbook‑scoped named range whose RefersTo formula points to a range in the linked workbook, and save the result.
    class Program
    {
        static void Main()
        {
            // Create the main workbook
            Workbook mainWb = new Workbook();

            // Add an external link to the external workbook file
            // The external file name can be a relative or absolute path
            string externalFile = "ExternalWorkbook.xlsx";
            string[] externalSheets = new string[] { "Sheet1" };
            int linkIndex = mainWb.Worksheets.ExternalLinks.Add(externalFile, externalSheets);
            ExternalLink extLink = mainWb.Worksheets.ExternalLinks[linkIndex];

            // (Optional) set the data source explicitly – useful if you need to change it later
            extLink.DataSource = externalFile;

            // Add a workbook‑scoped named range that refers to a range in the external workbook
            // SheetIndex = 0 (global) makes the name workbook‑scoped
            int nameIdx = mainWb.Worksheets.Names.Add("ExternalRange");
            Name externalName = mainWb.Worksheets.Names[nameIdx];
            // The RefersTo formula must start with '=' and include the external file in brackets
            externalName.RefersTo = "='[ExternalWorkbook.xlsx]Sheet1'!$A$1:$B$2";

            // Save the workbook containing the external named range
            mainWb.Save("WorkbookWithExternalNamedRange.xlsx");
        }
    }
}
