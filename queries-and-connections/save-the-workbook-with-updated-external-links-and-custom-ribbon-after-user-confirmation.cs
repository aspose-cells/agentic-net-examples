// Title: Conditionally save an Excel workbook with updated external links and custom Ribbon XML as macro‑enabled XLSM using Aspose.Cells for .NET
// AI Prompts: Inject custom Ribbon XML into a workbook, replace folder names in external link data sources, refresh linked formulas with temporary workbooks, then prompt the user and save the file as .xlsm. | Write a C# program that loads an .xlsx file, updates each ExternalLinkCollection path, calls UpdateLinkedDataSource, sets Workbook.RibbonXml, asks for y/n confirmation, and saves as a macro‑enabled workbook.
// Common Searches: Aspose.Cells C# update external link paths and refresh linked data source | How to add custom Ribbon XML to an Excel file with Aspose.Cells .NET | Save workbook as macro‑enabled XLSM after user confirmation using Aspose.Cells | Prompt user before saving Excel workbook in a C# console application Aspose.Cells | Replace folder name in external link data source Aspose.Cells example
// Tags: set workbook ribbonxml aspnet cells | replace external link datasource path aspnet cells | refresh linked data source using updatelinkeddatasource | save workbook as macro-enabled xlsm aspnet cells | console user confirmation before saving excel file

using System;
using System.IO;
using Aspose.Cells;

// Loads an existing .xlsx workbook, assigns custom Ribbon XML, updates each external link's data source path, creates temporary workbooks to refresh linked formulas, asks the user to confirm, and saves the result as a macro‑enabled .xlsm file.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Set custom Ribbon XML to customize the Excel UI
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
        workbook.RibbonXml = ribbonXml; // uses Workbook.RibbonXml property

        // Update external link data sources if any exist
        ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;
        for (int i = 0; i < externalLinks.Count; i++)
        {
            // Example modification: replace a folder name in the data source path
            string oldSource = externalLinks[i].DataSource;
            string newSource = oldSource.Replace("oldfolder", "newfolder");
            externalLinks[i].DataSource = newSource;
        }

        // Refresh external data by providing matching external workbooks
        if (externalLinks.Count > 0)
        {
            Workbook[] externalWorkbooks = new Workbook[externalLinks.Count];
            for (int i = 0; i < externalLinks.Count; i++)
            {
                Workbook ext = new Workbook();
                // Ensure FileName matches the DataSource file name (without path)
                ext.FileName = Path.GetFileName(externalLinks[i].DataSource);
                // Populate a sample value that linked formulas can retrieve
                ext.Worksheets[0].Cells["A1"].PutValue("Updated");
                externalWorkbooks[i] = ext;
            }
            workbook.UpdateLinkedDataSource(externalWorkbooks); // uses Workbook.UpdateLinkedDataSource method
        }

        // Ask the user for confirmation before saving
        Console.WriteLine("Do you want to save the workbook with updated links and custom ribbon? (y/n)");
        string answer = Console.ReadLine();
        if (!string.IsNullOrEmpty(answer) && answer.Trim().ToLower() == "y")
        {
            // Save as a macro‑enabled workbook to retain Ribbon XML
            workbook.Save("output.xlsm", SaveFormat.Xlsm); // uses Workbook.Save(string, SaveFormat)
            Console.WriteLine("Workbook saved as output.xlsm");
        }
        else
        {
            Console.WriteLine("Save operation cancelled by the user.");
        }

        // Release resources
        workbook.Dispose();
    }
}
