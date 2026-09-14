// Title: Replace <br> and <br/> tags with newline characters after loading HTML into an Aspose.Cells workbook (C#)
// AI Prompts: Load an HTML string into a Workbook using HtmlLoadOptions, then replace all <br> and <br/> tags with '\n' via Workbook.Replace before saving. | Show how to transform HTML line‑break tags into Excel cell line feeds in a C# Aspose.Cells example.
// Common Searches: C# Aspose.Cells replace HTML <br> tags with newline after importing HTML | How to convert <br/> tags to line feeds in an Excel workbook using Aspose.Cells | Aspose.Cells HtmlLoadOptions replace line break tags before saving workbook
// Tags: Workbook.Replace html line break conversion | Aspose.Cells load html with HtmlLoadOptions C# | convert html br tags to newline in Excel | replace html line breaks with cell line feeds Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// Loads an HTML string into a Workbook, uses Workbook.Replace to change <br> and <br/> tags to '\n', and saves the result as Result.xlsx.
class ReplaceBrWithLineFeed
{
    static void Main()
    {
        // Sample HTML containing <br> tags
        string html = "<p>First line<br>Second line<br/>Third line</p>";

        // Convert the HTML string to a memory stream
        byte[] htmlBytes = Encoding.UTF8.GetBytes(html);
        using (MemoryStream stream = new MemoryStream(htmlBytes))
        {
            // Load the HTML into a workbook (optional load options can be set)
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            loadOptions.SupportDivTag = true; // enable <div> support if needed
            Workbook workbook = new Workbook(stream, loadOptions);

            // Replace <br> and <br/> tags with line feed characters
            workbook.Replace("<br>", "\n");
            workbook.Replace("<br/>", "\n");

            // Save the resulting workbook
            workbook.Save("Result.xlsx");
        }
    }
}
