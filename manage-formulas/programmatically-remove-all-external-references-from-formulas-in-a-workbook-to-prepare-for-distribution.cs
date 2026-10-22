// Title: Programmatically strip external references from Excel formulas using Aspose.Cells in C#
// Common Searches: C# code to delete external links in an Excel file using Aspose.Cells | Aspose.Cells RemoveExternalLinks method example for .NET | How to clean workbook of external references before distribution with Aspose.Cells | Remove external workbook references from formulas in a .xlsx using C# | Strip external data source links from Excel workbook using Aspose.Cells API
// Tags: Aspose.Cells RemoveExternalLinks API usage | C# Excel workbook external link removal | clean Excel formulas external links .NET | prepare Excel file for distribution without external references | strip external workbook connections using Aspose.Cells

using System;
using Aspose.Cells;

// The sample loads an existing workbook, invokes the RemoveExternalLinks method to purge all external workbook, named range, and data source references from formulas, and then saves the cleaned workbook ready for safe distribution.
class Program
{
    static void Main()
    {
        // Load the workbook that contains external references.
        Workbook workbook = new Workbook("InputWorkbook.xlsx");

        // Remove all external links and references from formulas.
        // This clears any references to other workbooks, named ranges, or external data sources.
        workbook.RemoveExternalLinks();

        // Save the cleaned workbook ready for distribution.
        workbook.Save("OutputWorkbook.xlsx");
    }
}
