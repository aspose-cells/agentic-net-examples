using System;
using System.IO;
using Aspose.Cells;

class ExportTimeline
{
    static void Main()
    {
        string templatePath = "template.xltx";

        Workbook workbook;
        if (File.Exists(templatePath))
        {
            workbook = new Workbook(templatePath);
        }
        else
        {
            // If the template file is missing, create a new workbook as a fallback
            workbook = new Workbook();
        }

        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}