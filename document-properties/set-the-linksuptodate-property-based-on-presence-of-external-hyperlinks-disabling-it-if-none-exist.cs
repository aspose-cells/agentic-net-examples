using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class SetLinksUpToDateDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a sample hyperlink (remove/comment this block to test the "no hyperlink" case)
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Hyperlinks.Add("A1", 1, 1, "https://www.aspose.com");

        // Check if any worksheet contains at least one hyperlink
        bool hasHyperlink = false;
        foreach (Worksheet ws in workbook.Worksheets)
        {
            if (ws.Hyperlinks.Count > 0)
            {
                hasHyperlink = true;
                break;
            }
        }

        // Set the LinksUpToDate property based on the presence of hyperlinks
        BuiltInDocumentPropertyCollection builtInProps = workbook.BuiltInDocumentProperties;
        builtInProps.LinksUpToDate = hasHyperlink;

        // Save the workbook
        workbook.Save("LinksUpToDateResult.xlsx");
    }
}