// Title: Batch hide final worksheet in multiple Excel workbooks – Aspose.Cells C# example
// Description: A C# console program that loads a list of Excel files with Aspose.Cells, finds each workbook's last worksheet, hides it using SetVisible(false, true), and saves the changes back to the original location or a new folder.
// Keywords: Aspose.Cells hide worksheet | C# batch Excel processing | SetVisible false Aspose | hide final sheet multiple workbooks | Aspose.Cells Workbook.Save | bulk hide Excel sheet .NET
// Common Searches: hide final worksheet Aspose.Cells C# | batch hide Excel sheets .NET | process multiple workbooks hide sheet | Aspose.Cells SetVisible example | bulk Excel worksheet visibility C#
// Developer Intent: Iterate through a collection of Excel files, conceal each file's final worksheet, and persist the modification.
// Use Cases: Conceal a confidential summary tab across a fleet of financial reports before distribution. | Prepare template workbooks by automatically hiding configuration sheets for end‑users. | Perform compliance clean‑up by removing the final worksheet from archived spreadsheets in bulk.
// AI Prompts: Generate C# code that loops through a list of Excel file paths, hides the final worksheet with Aspose.Cells, and saves the workbooks. | Explain how SetVisible(false, true) hides a worksheet without raising errors in Aspose.Cells. | Adapt the example to write processed files to a separate output directory while keeping the originals unchanged.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace BatchHideLastWorksheet
{
    // A C# console program that loads a list of Excel files with Aspose.Cells, finds each workbook's last worksheet, hides it using SetVisible(false, true), and saves the changes back to the original location or a new folder.
    class Program
    {
        static void Main()
        {
            // List of workbook file paths to process
            List<string> workbookFiles = new List<string>
            {
                "Book1.xlsx",
                "Book2.xlsx",
                "Book3.xlsx"
                // Add more file paths as needed
            };

            foreach (string filePath in workbookFiles)
            {
                // Load the workbook (uses the provided Workbook(string) constructor)
                Workbook workbook = new Workbook(filePath);

                // Determine the index of the last worksheet
                int lastIndex = workbook.Worksheets.Count - 1;

                if (lastIndex >= 0)
                {
                    // Hide the last worksheet using the SetVisible method (provided rule)
                    // isVisible = false to hide, ignoreError = true to suppress errors
                    workbook.Worksheets[lastIndex].SetVisible(false, true);
                }

                // Save the modified workbook (uses the provided Save(string) method)
                // Overwrites the original file; change the path if a separate output is desired
                workbook.Save(filePath);
            }

            Console.WriteLine("Processing completed. Last worksheets have been hidden.");
        }
    }
}
