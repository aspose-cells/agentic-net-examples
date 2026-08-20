// Title: Set identical left, center, and right page headers on every worksheet with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, defines header strings (file name, page number, date), iterates through all worksheets, uses PageSetup.SetHeader to assign left, center, and right sections, and saves the updated file.
// Keywords: Aspose.Cells C# set header all worksheets | page header batch Aspose.Cells | Excel workbook header automation .NET | SetHeader multiple sheets | apply page header Aspose.Cells | C# Excel header script
// Common Searches: C# Aspose.Cells set same header for all sheets | How to apply page header to every worksheet in Excel using Aspose | Batch update Excel headers with Aspose.Cells .NET | Set left center right header programmatically Aspose.Cells | Apply file name and date header to all worksheets
// Developer Intent: The developer wants to programmatically apply identical left, center, and right page headers to every worksheet in an Excel workbook.
// Use Cases: Standardize printed reports by adding file name, page numbers, and date to each sheet's header. | Prepare multi‑sheet financial workbooks for printing with consistent branding across all worksheets. | Automate header insertion for large batches of generated workbooks before distribution.
// AI Prompts: Generate C# code using Aspose.Cells to set custom left, center, and right headers on all worksheets and save the workbook. | Show how to customize header text dynamically based on each worksheet's name while iterating with Aspose.Cells. | Provide best‑practice error handling for loading a large workbook and applying page headers with Aspose.Cells.

using System;
using Aspose.Cells;

namespace BatchHeaderSetter
{
    // Loads an Excel workbook, defines header strings (file name, page number, date), iterates through all worksheets, uses PageSetup.SetHeader to assign left, center, and right sections, and saves the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths (adjust as needed)
            string inputPath = "LargeWorkbook.xlsx";
            string outputPath = "LargeWorkbook_WithHeaders.xlsx";

            // Load the existing workbook (uses the provided Workbook constructor)
            Workbook workbook = new Workbook(inputPath);

            // Define the header script to be applied to all worksheets
            // Example: left section = file name, center = page number, right = date
            string leftHeader = "&F";
            string centerHeader = "Page &P of &N";
            string rightHeader = "&D";

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the PageSetup object of the current worksheet
                PageSetup pageSetup = sheet.PageSetup;

                // Apply the same header to all three sections
                pageSetup.SetHeader(0, leftHeader);   // Left section
                pageSetup.SetHeader(1, centerHeader); // Center section
                pageSetup.SetHeader(2, rightHeader);  // Right section
            }

            // Save the modified workbook (uses the provided Save method)
            workbook.Save(outputPath);
        }
    }
}
