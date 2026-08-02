// Title: Remove External Links from Excel Formulas Using Aspose.Cells (.NET)
// Description: C# code that loads an .xlsx workbook with Aspose.Cells, clears all external workbook references via ExternalLinkCollection, rewrites formulas to local cells, and saves a distribution‑ready file.
// Keywords: Aspose.Cells external link removal | C# clear Excel external references | ExternalLinkCollection Clear updateReferencesAsLocal | convert external formulas to local | Excel workbook distribution without links | .NET Excel API remove external links | strip external data sources from workbook | Aspose.Cells formula cleanup
// Common Searches: how to delete external workbook links with Aspose.Cells | C# example for removing Excel external references | Aspose.Cells clear external links before sharing file | update formulas to internal references .NET | batch process Excel files to eliminate external links
// Developer Intent: Eliminate every external workbook reference and adjust formulas to point to the current file.
// Use Cases: Prepare a financial model for client delivery by stripping hidden data connections. | Publish a template that must not contain links to other spreadsheets. | Run an automated job that sanitizes a folder of Excel files for compliance.
// AI Prompts: Generate C# code that removes all external links from an Excel workbook with Aspose.Cells and updates formulas to local references. | Explain the effect of ExternalLinkCollection.Clear(updateReferencesAsLocal:true) on formulas that originally pointed to other workbooks. | Provide a step‑by‑step tutorial for loading a workbook, clearing external references, and saving the cleaned version using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace RemoveExternalLinksDemo
{
    // C# code that loads an .xlsx workbook with Aspose.Cells, clears all external workbook references via ExternalLinkCollection, rewrites formulas to local cells, and saves a distribution‑ready file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook that may contain external links
            string inputPath = "input.xlsx";

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the collection of external links in the workbook
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

            // Remove all external links and update formulas to refer to the current workbook where possible
            externalLinks.Clear(updateReferencesAsLocal: true);

            // Save the cleaned workbook ready for distribution
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"External links removed and workbook saved to '{outputPath}'.");
        }
    }
}
