// Title: Remove External Links from Excel Formulas with Aspose.Cells for .NET (C#)
// Description: Load an Excel workbook, clear its ExternalLinkCollection, optionally update formulas to reference the current file, and save the sanitized workbook—ideal for preparing files for distribution.
// Keywords: Aspose.Cells external links | clear external references | remove workbook links C# | update formulas after link removal | Aspose.Cells ExternalLinkCollection | Excel sanitization .NET | remove #REF! external references | C# Excel automation | prepare workbook for distribution
// Common Searches: Aspose.Cells remove external links C# | clear external references in Excel with .NET | how to delete external links using Aspose.Cells | update formulas after removing external links | programmatically strip external references from workbook
// Developer Intent: Eliminate all external references in workbook formulas to create a clean file ready for sharing or publishing.
// Use Cases: Sanitize client‑facing workbooks by stripping links to source data while keeping internal calculations intact. | Automate batch processing of multiple Excel files to ensure no external references remain before archiving. | Prepare reusable templates that must not contain links to external files or sheets.
// AI Prompts: Write C# code using Aspose.Cells that loads an Excel file, clears all external links with formula updates, and saves the result. | Show how to iterate over a directory of .xlsx files, remove external references from each workbook using Aspose.Cells, and log any formulas that become #REF!. | Explain the effect of ExternalLinkCollection.Clear(true) on formulas that reference missing worksheets.

using System;
using Aspose.Cells;

namespace RemoveExternalLinksDemo
{
    // Load an Excel workbook, clear its ExternalLinkCollection, optionally update formulas to reference the current file, and save the sanitized workbook—ideal for preparing files for distribution.
    class Program
    {
        static void Main(string[] args)
        {
            // Load the workbook that may contain external links
            // Replace "input.xlsx" with the path to your source file
            Workbook workbook = new Workbook("input.xlsx");

            // Get the collection of external links in the workbook
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

            // Remove all external links.
            // Pass true to update references in formulas to refer to the current workbook where possible.
            // If a referenced sheet does not exist locally, the formula will become #REF!.
            externalLinks.Clear(true);

            // Save the cleaned workbook.
            // Replace "output.xlsx" with the desired output path.
            workbook.Save("output.xlsx");

            Console.WriteLine("External links removed and workbook saved successfully.");
        }
    }
}
