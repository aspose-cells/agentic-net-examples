using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class CheckHiddenExternalLinks
    {
        public static void Run()
        {
            // Load an existing XLSX workbook (default format)
            Workbook workbook = new Workbook("input.xlsx");

            // Ensure the workbook contains external links
            if (workbook.Worksheets.ExternalLinks.Count == 0)
            {
                Console.WriteLine("The workbook does not contain any external links.");
                return;
            }

            // Iterate through all external links and display their visibility status
            Console.WriteLine("External links in the workbook:");
            foreach (ExternalLink link in workbook.Worksheets.ExternalLinks)
            {
                // IsVisible indicates whether the external link is visible in Excel
                Console.WriteLine($"DataSource: {link.DataSource}");
                Console.WriteLine($"IsVisible : {link.IsVisible}");
                Console.WriteLine(new string('-', 40));
            }

            // Example: Identify hidden external links (IsVisible == false)
            Console.WriteLine("Hidden external links:");
            bool anyHidden = false;
            foreach (ExternalLink link in workbook.Worksheets.ExternalLinks)
            {
                if (!link.IsVisible)
                {
                    anyHidden = true;
                    Console.WriteLine($"Hidden Link DataSource: {link.DataSource}");
                }
            }

            if (!anyHidden)
            {
                Console.WriteLine("No hidden external links were found.");
            }

            // Optional: Save the workbook after any modifications (if needed)
            // workbook.Save("output.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            CheckHiddenExternalLinks.Run();
        }
    }
}