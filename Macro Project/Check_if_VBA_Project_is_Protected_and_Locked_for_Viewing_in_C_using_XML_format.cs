using System;
using Aspose.Cells;
using Aspose.Cells.Vba;
using System.Xml.Linq;

class CheckVbaProtection
{
    static void Main()
    {
        // Path to the macro-enabled Excel file
        string filePath = "sample.xlsm";

        // Load the workbook
        Workbook workbook = new Workbook(filePath);

        // Get the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Check protection flags
        bool isProtected = vbaProject.IsProtected;
        bool isLockedForViewing = vbaProject.IslockedForViewing;

        // Build an XML document containing the results
        XDocument resultXml = new XDocument(
            new XElement("VbaProjectProtection",
                new XElement("IsProtected", isProtected),
                new XElement("IsLockedForViewing", isLockedForViewing)
            )
        );

        // Output the XML to the console
        Console.WriteLine(resultXml);
    }
}