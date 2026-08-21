// Title: Remove hidden external link from an Excel workbook using Aspose.Cells for .NET
// Description: Load a workbook, scan the Worksheets.ExternalLinks collection for links with IsVisible = false, delete the hidden link, verify that no invisible links remain, and save the updated file.
// Keywords: Aspose.Cells remove hidden external link | delete invisible external link .NET | ExternalLinkCollection IsVisible | C# Aspose.Cells external links | verify external link removal
// Common Searches: how to delete hidden external links with Aspose.Cells | remove invisible external link from Excel using C# | check for hidden external links after removal Aspose | Aspose.Cells external link visibility
// Developer Intent: Programmatically locate and delete a non‑visible external link in a workbook and confirm its successful removal.
// Use Cases: Sanitize workbooks before distribution to eliminate hidden data connections. | Strip confidential external sources that were concealed in shared spreadsheets. | Ensure compliance by confirming only visible external links exist after processing.
// AI Prompts: Write C# code with Aspose.Cells that removes all ExternalLink objects where IsVisible is false and returns a boolean indicating remaining hidden links. | Create a method that logs the name and index of each hidden external link before deletion using Aspose.Cells for .NET. | Explain step‑by‑step how to verify that no hidden external links are left in a workbook after removal.

using System;
using Aspose.Cells;

// Load a workbook, scan the Worksheets.ExternalLinks collection for links with IsVisible = false, delete the hidden link, verify that no invisible links remain, and save the updated file.
class RemoveHiddenExternalLink
{
    static void Main()
    {
        // Load the workbook containing external links
        Workbook workbook = new Workbook("input.xlsx");

        // Access the external links collection
        ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

        // Locate the index of the hidden external link (IsVisible == false)
        int hiddenLinkIndex = -1;
        for (int i = 0; i < externalLinks.Count; i++)
        {
            if (!externalLinks[i].IsVisible)
            {
                hiddenLinkIndex = i;
                break;
            }
        }

        // Remove the hidden external link if it exists
        if (hiddenLinkIndex >= 0)
        {
            externalLinks.RemoveAt(hiddenLinkIndex);
            Console.WriteLine("Hidden external link removed.");
        }
        else
        {
            Console.WriteLine("No hidden external link found.");
        }

        // Verify that no hidden external links remain
        bool hiddenLinkStillExists = false;
        for (int i = 0; i < externalLinks.Count; i++)
        {
            if (!externalLinks[i].IsVisible)
            {
                hiddenLinkStillExists = true;
                break;
            }
        }
        Console.WriteLine("Verification - hidden link present after removal: " + hiddenLinkStillExists);

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
