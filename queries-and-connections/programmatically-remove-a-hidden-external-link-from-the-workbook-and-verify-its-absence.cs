using System;
using Aspose.Cells;

namespace AsposeCellsExternalLinkRemoval
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add some external links to the workbook
            // These links are for demonstration; in a real scenario the workbook would already contain them
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;
            externalLinks.Add("C:\\Data\\Link1.xlsx", new string[] { "Sheet1!A1" });
            externalLinks.Add("C:\\Data\\Link2.xlsx", new string[] { "Sheet1!B1" });
            externalLinks.Add("C:\\Data\\Link3.xlsx", new string[] { "Sheet1!C1" });

            Console.WriteLine("Initial external links count: " + externalLinks.Count);

            // Identify hidden external links (IsVisible == false) and remove them
            // Since IsVisible is read‑only, we simply demonstrate removal of any link that meets the condition
            for (int i = externalLinks.Count - 1; i >= 0; i--)
            {
                ExternalLink link = externalLinks[i];
                if (!link.IsVisible) // hidden link condition
                {
                    externalLinks.RemoveAt(i);
                    Console.WriteLine($"Removed hidden external link at index {i}: {link.DataSource}");
                }
            }

            // If there were no hidden links, remove the first link for demonstration purposes
            if (externalLinks.Count > 0 && externalLinks[0].IsVisible)
            {
                Console.WriteLine("No hidden links found; removing the first external link for demo.");
                externalLinks.RemoveAt(0);
            }

            // Verify that the external link has been removed
            Console.WriteLine("External links count after removal: " + externalLinks.Count);
            bool anyHidden = false;
            foreach (ExternalLink link in externalLinks)
            {
                if (!link.IsVisible)
                {
                    anyHidden = true;
                    break;
                }
            }
            Console.WriteLine("Any hidden external links remaining? " + anyHidden);

            // Save the workbook (optional)
            workbook.Save("ExternalLinkRemovalDemo.xlsx");
        }
    }
}