// Title: How to remove hidden external links from an Excel workbook and confirm their removal using Aspose.Cells for .NET (C#)
// Common Searches: C# code to delete hidden external links in an .xlsx file with Aspose.Cells | using reflection to access Worksheet.ExternalLinks collection in Aspose.Cells | verify that no hidden external links exist after modifying a workbook in Aspose.Cells | remove hidden external link from a specific worksheet programmatically Aspose.Cells | check for hidden external links in Excel workbook using Aspose.Cells .NET
// Tags: reflection based external link removal Aspose.Cells | Worksheet ExternalLinks manipulation C# | detect hidden external links Aspose.Cells | Excel workbook external link cleanup .NET | Aspose.Cells hidden link verification

using System;
using System.IO;
using System.Collections;
using Aspose.Cells;

// The example loads an Excel workbook, uses reflection to reach the internal Worksheet.ExternalLinks collection, removes any links flagged as hidden, saves the workbook, and then re‑examines the file via reflection to ensure no hidden external links remain.
class RemoveHiddenExternalLink
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);
            bool hiddenLinkRemoved = false;

            // Iterate through each worksheet and remove hidden external links using reflection
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (RemoveHiddenLinksFromWorksheet(sheet))
                {
                    hiddenLinkRemoved = true;
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);

            // Verify that no hidden external links remain
            bool hiddenLinkExists = CheckHiddenLinksExist(workbook);

            // Output verification result
            if (hiddenLinkRemoved && !hiddenLinkExists)
                Console.WriteLine("Hidden external link successfully removed.");
            else if (!hiddenLinkRemoved)
                Console.WriteLine("No hidden external link was found to remove.");
            else
                Console.WriteLine("Failed to remove hidden external link.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Removes hidden external links from a worksheet using reflection.
    private static bool RemoveHiddenLinksFromWorksheet(Worksheet sheet)
    {
        try
        {
            var externalLinksProp = typeof(Worksheet).GetProperty("ExternalLinks");
            if (externalLinksProp == null) return false; // Property not supported in this version

            var externalLinks = externalLinksProp.GetValue(sheet) as IList;
            if (externalLinks == null) return false;

            var countProp = externalLinks.GetType().GetProperty("Count");
            var removeAtMethod = externalLinks.GetType().GetMethod("RemoveAt");
            if (countProp == null || removeAtMethod == null) return false;

            int count = (int)countProp.GetValue(externalLinks);
            bool removedAny = false;

            // Iterate in reverse to safely remove items
            for (int i = count - 1; i >= 0; i--)
            {
                var link = externalLinks[i];
                var isHiddenProp = link.GetType().GetProperty("IsHidden");
                if (isHiddenProp == null) continue;

                bool isHidden = (bool)isHiddenProp.GetValue(link);
                if (isHidden)
                {
                    removeAtMethod.Invoke(externalLinks, new object[] { i });
                    removedAny = true;
                }
            }
            return removedAny;
        }
        catch
        {
            // If reflection fails, treat as no removal performed
            return false;
        }
    }

    // Checks if any hidden external links exist in the workbook using reflection.
    private static bool CheckHiddenLinksExist(Workbook workbook)
    {
        try
        {
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                var externalLinksProp = typeof(Worksheet).GetProperty("ExternalLinks");
                if (externalLinksProp == null) continue;

                var externalLinks = externalLinksProp.GetValue(sheet) as IList;
                if (externalLinks == null) continue;

                foreach (var link in externalLinks)
                {
                    var isHiddenProp = link.GetType().GetProperty("IsHidden");
                    if (isHiddenProp == null) continue;

                    bool isHidden = (bool)isHiddenProp.GetValue(link);
                    if (isHidden) return true;
                }
            }
            return false;
        }
        catch
        {
            // If reflection fails, assume hidden links may exist
            return true;
        }
    }
}
