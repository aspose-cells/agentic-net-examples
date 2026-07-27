using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsHtmlCleanup
{
    // Author: Aspose.Cells .NET example
    class Program
    {
        static void Main()
        {
            // Sample HTML containing <br> tags with redundant spaces
            string html = "<p>Line1<br>   <br>Line2<br>   </p>";

            // Configure load options to delete redundant spaces after <br> tags
            HtmlLoadOptions loadOptions = new HtmlLoadOptions
            {
                DeleteRedundantSpaces = true // Removes unnecessary spaces around <br>
            };

            // Load the HTML into a workbook using the configured options
            byte[] htmlBytes = Encoding.UTF8.GetBytes(html);
            using (MemoryStream stream = new MemoryStream(htmlBytes))
            {
                Workbook workbook = new Workbook(stream, loadOptions);

                // Optionally, manipulate the workbook here

                // Save the workbook back to HTML; default save options preserve the cleaned text
                HtmlSaveOptions saveOptions = new HtmlSaveOptions();

                // Output file path
                string outputPath = "cleaned_output.html";
                workbook.Save(outputPath, saveOptions);

                Console.WriteLine($"HTML saved without redundant spaces after <br> tags to: {outputPath}");
            }
        }
    }
}