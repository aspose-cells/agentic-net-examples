using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and add some formatted data
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Hello");
        worksheet.Cells["A1"].GetStyle().Font.IsBold = true;
        worksheet.Cells["A2"].PutValue("World");
        worksheet.Cells["A2"].GetStyle().Font.Color = Color.Blue;

        // Configure HTML save options to disable embedded CSS (use only inline styles)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.DisableCss = true;                     // No CSS blocks will be generated
        htmlOptions.ExportWorksheetCSSSeparately = false; // Ensure no separate CSS files are created

        string htmlFile = "output.html";
        workbook.Save(htmlFile, htmlOptions);               // Save the workbook as HTML

        // Create an external stylesheet with custom styles
        string cssFile = "custom.css";
        string cssContent = @"
            body {
                font-family: Arial, sans-serif;
                background-color: #f9f9f9;
            }
            td {
                border: 1px solid #ccc;
                padding: 5px;
            }
        ";
        File.WriteAllText(cssFile, cssContent);

        // Insert a <link> tag for the external stylesheet into the generated HTML head section
        string html = File.ReadAllText(htmlFile);
        int headCloseIndex = html.IndexOf("</head>", StringComparison.OrdinalIgnoreCase);
        if (headCloseIndex >= 0)
        {
            string linkTag = $@"<link rel=""stylesheet"" type=""text/css"" href=""{cssFile}"">" + Environment.NewLine;
            html = html.Insert(headCloseIndex, linkTag);
            File.WriteAllText(htmlFile, html);
        }
    }
}