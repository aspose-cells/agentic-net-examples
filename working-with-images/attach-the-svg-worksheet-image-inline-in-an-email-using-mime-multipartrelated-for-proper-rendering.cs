// Title: Create a .eml email with an Excel worksheet rendered as SVG and embedded inline using MIME multipart/related in C# (Aspose.Cells)
// Common Searches: C# generate an .eml file that displays an Excel sheet as inline SVG using Aspose.Cells | how to embed a rendered worksheet SVG into an email body with MIME multipart/related in .NET | Aspose.Cells render worksheet to SVG and attach as inline image in email message | build MIME multipart/related email with base64 encoded SVG from Excel in C# | send Excel worksheet as SVG image inside HTML email using Aspose.Cells
// Tags: Aspose.Cells SVG sheet rendering | C# MIME multipart/related construction | email inline SVG embedding | base64 encode SVG for email | Content-ID image/svg+xml attachment

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// The example loads or creates an Excel workbook, uses Aspose.Cells to render the first worksheet to SVG, builds a MIME multipart/related message with an HTML part referencing the SVG via a Content-ID, encodes the SVG as base64, and writes the complete message to a .eml file for sending.
class Program
{
    static void Main()
    {
        const string workbookPath = "input.xlsx";
        Workbook workbook = null;

        // Load or create workbook
        try
        {
            if (File.Exists(workbookPath))
            {
                workbook = new Workbook(workbookPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Name = "Sample";

                ws.Cells["A1"].PutValue("Product");
                ws.Cells["B1"].PutValue("Price");
                ws.Cells["A2"].PutValue("Apple");
                ws.Cells["B2"].PutValue(1.20);
                ws.Cells["A3"].PutValue("Banana");
                ws.Cells["B3"].PutValue(0.80);

                workbook.Save(workbookPath);
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Failed to load or create workbook: {ex.Message}");
            return;
        }

        // Render first worksheet to SVG
        string svgContent;
        try
        {
            ImageOrPrintOptions printOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Svg, // Aspose.Cells.Drawing.ImageType
                OnePagePerSheet = true
            };

            SheetRender renderer = new SheetRender(workbook.Worksheets[0], printOptions);
            using (MemoryStream svgStream = new MemoryStream())
            {
                renderer.ToImage(0, svgStream);
                svgStream.Position = 0;
                using (StreamReader reader = new StreamReader(svgStream, Encoding.UTF8))
                {
                    svgContent = reader.ReadToEnd();
                }
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"SVG rendering failed: {ex.Message}");
            return;
        }

        if (string.IsNullOrWhiteSpace(svgContent))
        {
            Console.Error.WriteLine("Failed to generate SVG content.");
            return;
        }

        // Build MIME multipart/related message
        string boundary = "----=_Part_" + Guid.NewGuid().ToString("N");

        string htmlBody = @"<html>
<head>
    <meta charset=""UTF-8"">
    <title>Worksheet SVG Inline</title>
</head>
<body>
    <h2>Worksheet rendered as SVG</h2>
    <img src=""cid:svgImage"" alt=""Worksheet SVG""/>
</body>
</html>";

        string svgBase64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(svgContent));

        StringBuilder mimeBuilder = new StringBuilder();

        mimeBuilder.AppendLine("From: sender@example.com");
        mimeBuilder.AppendLine("To: recipient@example.com");
        mimeBuilder.AppendLine("Subject: Worksheet SVG Inline");
        mimeBuilder.AppendLine("MIME-Version: 1.0");
        mimeBuilder.AppendLine($"Content-Type: multipart/related; boundary=\"{boundary}\"");
        mimeBuilder.AppendLine();

        // HTML part
        mimeBuilder.AppendLine($"--{boundary}");
        mimeBuilder.AppendLine("Content-Type: text/html; charset=\"utf-8\"");
        mimeBuilder.AppendLine("Content-Transfer-Encoding: 7bit");
        mimeBuilder.AppendLine();
        mimeBuilder.AppendLine(htmlBody);
        mimeBuilder.AppendLine();

        // SVG part
        mimeBuilder.AppendLine($"--{boundary}");
        mimeBuilder.AppendLine("Content-Type: image/svg+xml; name=\"sheet.svg\"");
        mimeBuilder.AppendLine("Content-Transfer-Encoding: base64");
        mimeBuilder.AppendLine("Content-ID: <svgImage>");
        mimeBuilder.AppendLine("Content-Disposition: inline; filename=\"sheet.svg\"");
        mimeBuilder.AppendLine();

        for (int i = 0; i < svgBase64.Length; i += 76)
        {
            int chunkSize = Math.Min(76, svgBase64.Length - i);
            mimeBuilder.AppendLine(svgBase64.Substring(i, chunkSize));
        }
        mimeBuilder.AppendLine();

        // Closing boundary
        mimeBuilder.AppendLine($"--{boundary}--");
        mimeBuilder.AppendLine();

        // Write MIME message to .eml file
        const string emlPath = "worksheet_svg_email.eml";
        try
        {
            File.WriteAllText(emlPath, mimeBuilder.ToString(), Encoding.UTF8);
            Console.WriteLine($"MIME message written to '{emlPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Failed to write EML file: {ex.Message}");
        }
    }
}
