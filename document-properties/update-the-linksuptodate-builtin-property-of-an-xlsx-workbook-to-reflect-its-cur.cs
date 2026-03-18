using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsLinkStatusDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Determine if the workbook contains any hyperlinks
            bool hasHyperlinks = false;
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.Hyperlinks.Count > 0)
                {
                    hasHyperlinks = true;
                    break;
                }
            }

            // Update the built‑in property to reflect the current hyperlink status
            BuiltInDocumentPropertyCollection builtInProps = workbook.BuiltInDocumentProperties;
            builtInProps.LinksUpToDate = hasHyperlinks; // true if hyperlinks exist, otherwise false

            // Optionally, add a hyperlink to demonstrate the property being true
            if (!hasHyperlinks)
            {
                Worksheet firstSheet = workbook.Worksheets[0];
                // Add a hyperlink at cell A1 pointing to Aspose website
                firstSheet.Hyperlinks.Add("A1", 1, 1, "https://www.aspose.com");
                // After adding, set the property to true
                builtInProps.LinksUpToDate = true;
            }

            // Save the workbook with the updated property
            workbook.Save("output.xlsx");
        }
    }
}