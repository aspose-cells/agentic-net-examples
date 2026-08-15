// Title: C# – Update External Links, Add Custom Ribbon UI, and Save Workbook with Confirmation using Aspose.Cells
// Description: A C# console example that checks for a main and an external workbook, loads them with Aspose.Cells, refreshes the main workbook’s external links via UpdateLinkedDataSource, injects a custom Ribbon XML definition, asks the user to confirm, and saves the result as a macro‑enabled XLSM file.
// Keywords: Aspose.Cells C# | update external links Excel | UpdateLinkedDataSource | custom RibbonXml | embed custom ribbon Aspose.Cells | save workbook after prompt | macro enabled XLSM | .NET Excel automation | conditional workbook save | Excel linked data source refresh
// Common Searches: Aspose.Cells refresh external links C# | How to add custom Ribbon XML with Aspose.Cells | Save Excel file after user confirmation .NET | Update linked data source and preserve macros Aspose.Cells | C# code to embed Ribbon UI in XLSM
// Developer Intent: Refresh linked data, embed a custom Ribbon UI, and save the workbook only when the user approves.
// Use Cases: Ensure all external references are up‑to‑date before distributing a report. | Add a company‑branded tab and button to the Excel Ribbon for quick macro access. | Prevent accidental overwrites by prompting the user before saving changes.
// AI Prompts: Write C# code using Aspose.Cells that updates multiple external links, sets RibbonXml, and saves the workbook as XLSM after a yes/no prompt. | Explain the RibbonXml property in Aspose.Cells and demonstrate how to embed a custom UI definition into an Excel file. | Show how to handle missing external workbook files gracefully when calling UpdateLinkedDataSource with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// A C# console example that checks for a main and an external workbook, loads them with Aspose.Cells, refreshes the main workbook’s external links via UpdateLinkedDataSource, injects a custom Ribbon XML definition, asks the user to confirm, and saves the result as a macro‑enabled XLSM file.
class Program
{
    static void Main()
    {
        try
        {
            const string mainPath = "Main.xlsx";
            const string externalPath = "External.xlsx";

            // Verify that the required files exist
            if (!File.Exists(mainPath))
            {
                Console.WriteLine($"Error: File '{mainPath}' not found.");
                return;
            }

            if (!File.Exists(externalPath))
            {
                Console.WriteLine($"Error: File '{externalPath}' not found.");
                return;
            }

            // Load the main workbook that may contain external links
            using (Workbook mainWorkbook = new Workbook(mainPath))
            {
                // Load an external workbook that is referenced by the main workbook
                using (Workbook externalWorkbook = new Workbook(externalPath))
                {
                    // Update the external links in the main workbook with the latest data
                    mainWorkbook.UpdateLinkedDataSource(new Workbook[] { externalWorkbook });
                }

                // Define custom Ribbon XML
                string ribbonXml =
                    "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
                    "  <ribbon>" +
                    "    <tabs>" +
                    "      <tab id=\"customTab\" label=\"My Tab\">" +
                    "        <group id=\"customGroup\" label=\"My Group\">" +
                    "          <button id=\"customButton\" label=\"My Button\" size=\"large\" />" +
                    "        </group>" +
                    "      </tab>" +
                    "    </tabs>" +
                    "  </ribbon>" +
                    "</customUI>";

                // Apply the custom Ribbon UI to the workbook
                mainWorkbook.RibbonXml = ribbonXml;

                // Ask the user for confirmation before saving
                Console.Write("Save the workbook with updated links and custom ribbon? (y/n): ");
                string answer = Console.ReadLine();

                if (!string.IsNullOrEmpty(answer) && answer.Trim().Equals("y", StringComparison.OrdinalIgnoreCase))
                {
                    // Save the workbook (XLSM to preserve macros and Ribbon UI)
                    mainWorkbook.Save("Main_Updated.xlsm", SaveFormat.Xlsm);
                    Console.WriteLine("Workbook saved successfully.");
                }
                else
                {
                    Console.WriteLine("Save operation cancelled by the user.");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
