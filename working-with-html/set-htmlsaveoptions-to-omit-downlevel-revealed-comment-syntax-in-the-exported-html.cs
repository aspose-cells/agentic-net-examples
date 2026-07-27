using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (replace with the provided create rule if one exists)
        Workbook workbook = new Workbook();

        // Add sample data (optional)
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample text");

        // Configure HTML save options to omit downlevel‑revealed conditional comments
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            DisableDownlevelRevealedComments = true
        };

        // Save the workbook as HTML (replace with the provided save rule if one exists)
        workbook.Save("output.html", htmlOptions);
    }
}

// Author: Aspose.Cells .NET example (comments follow C# conventions)