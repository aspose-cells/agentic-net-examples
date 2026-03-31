using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class DetectHiddenExternalHyperlinks
    {
        public static void Run()
        {
            // Path to the XLSX file to be examined
            string filePath = "input.xlsx";

            // Load the workbook (default format is inferred from the file extension)
            Workbook workbook = new Workbook(filePath);

            // Access the collection of external links in the workbook
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

            // If there are no external links, inform the user and exit
            if (externalLinks.Count == 0)
            {
                Console.WriteLine("The workbook does not contain any external links.");
                return;
            }

            // Iterate through each external link and report those that are hidden (IsVisible == false)
            Console.WriteLine("Hidden external hyperlinks found in the workbook:");
            bool hiddenFound = false;

            for (int i = 0; i < externalLinks.Count; i++)
            {
                ExternalLink link = externalLinks[i];

                // The IsVisible property indicates whether the external link is visible in Excel.
                // A value of false means the link is hidden.
                if (!link.IsVisible)
                {
                    hiddenFound = true;
                    Console.WriteLine($"- Link Index: {i}");
                    Console.WriteLine($"  Data Source : {link.DataSource}");
                    Console.WriteLine($"  Type        : {link.Type}");
                    Console.WriteLine($"  Path Type   : {link.PathType}");
                }
            }

            if (!hiddenFound)
            {
                Console.WriteLine("No hidden external hyperlinks were detected.");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            DetectHiddenExternalHyperlinks.Run();
        }
    }
}