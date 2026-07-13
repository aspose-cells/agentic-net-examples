using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Author: Aspose.Cells .NET example – removes redundant spaces after line breaks when converting HTML to Excel and back to HTML
    class RemoveRedundantSpacesDemo
    {
        static void Main()
        {
            // Sample HTML containing redundant spaces and line breaks (<br>)
            string html = "<p>   This    text   has   redundant   spaces   <br>   and   more   spaces   </p>";

            // Configure HtmlLoadOptions to delete redundant spaces during import
            HtmlLoadOptions loadOptions = new HtmlLoadOptions
            {
                DeleteRedundantSpaces = true // removes extra spaces after <br> tags
            };

            // Convert the HTML string to a byte array and load it into a workbook using the options above
            byte[] htmlBytes = Encoding.UTF8.GetBytes(html);
            using (MemoryStream stream = new MemoryStream(htmlBytes))
            {
                Workbook workbook = new Workbook(stream, loadOptions);
                Worksheet sheet = workbook.Worksheets[0];

                // Optional: display the cleaned cell text in console
                Console.WriteLine("Cell A1 after removing redundant spaces:");
                Console.WriteLine(sheet.Cells["A1"].StringValue);

                // Save the workbook back to HTML – the cell text will no longer contain redundant spaces
                workbook.Save("CleanedOutput.html", SaveFormat.Html);
            }
        }
    }
}