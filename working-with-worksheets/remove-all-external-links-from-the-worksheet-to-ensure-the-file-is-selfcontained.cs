// Title: Remove external links from an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# code to strip external links from an Excel file using Aspose.Cells | Aspose.Cells .NET remove external references before saving workbook | Make an Excel workbook self‑contained by deleting external links in C# | How to use Workbook.RemoveExternalLinks method in Aspose.Cells | Remove external hyperlinks from .xlsx with Aspose.Cells C#
// Tags: Workbook.RemoveExternalLinks API | Aspose.Cells self‑contained workbook | C# eliminate external references Excel | clean Excel external links .NET | strip external hyperlinks Aspose.Cells

using Aspose.Cells;
using System;

// The example loads an Excel file (input.xlsx), calls Workbook.RemoveExternalLinks() to delete all external links, and saves the resulting self‑contained workbook as output.xlsx.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Remove all external links so the file becomes self‑contained
        workbook.RemoveExternalLinks();

        // Save the cleaned workbook
        workbook.Save("output.xlsx");
    }
}
