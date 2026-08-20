// Title: C# – Aspose.Cells: Remove All Horizontal and Vertical Page Breaks from a Worksheet
// Description: Demonstrates how to clear every manual page break in an Aspose.Cells worksheet using C#. The code creates or loads a workbook, empties the HorizontalPageBreaks and VerticalPageBreaks collections, and saves the file so Excel applies its default automatic pagination.
// Keywords: Aspose.Cells remove page breaks C# | clear worksheet page breaks .NET | delete manual page breaks Aspose | automatic pagination Excel C# | horizontal vertical page break removal | Aspose.Cells printing layout | C# Excel pagination reset
// Common Searches: how to clear all page breaks in Aspose.Cells C# | remove manual page breaks from Excel worksheet using .NET | Aspose.Cells automatic pagination after deleting page breaks | C# code to clear horizontal and vertical page breaks in Excel file | reset pagination in Aspose.Cells workbook
// Developer Intent: Delete every existing horizontal and vertical page break in a worksheet so the document uses Excel’s built‑in automatic pagination.
// Use Cases: Prepare a report for printing with default page layout after removing custom breaks. | Sanitize a shared template to ensure consistent pagination for all users. | Reset pagination in dynamically generated workbooks before final distribution.
// AI Prompts: Write C# code with Aspose.Cells that removes all page breaks from every worksheet in a workbook and then saves it. | Explain how Aspose.Cells recalculates automatic pagination once manual page breaks are cleared. | Show an example that clears page breaks and configures print options such as fit‑to‑page in Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to clear every manual page break in an Aspose.Cells worksheet using C#. The code creates or loads a workbook, empties the HorizontalPageBreaks and VerticalPageBreaks collections, and saves the file so Excel applies its default automatic pagination.
class RemovePageBreaksDemo
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Remove all horizontal page breaks
        worksheet.HorizontalPageBreaks.Clear();

        // Remove all vertical page breaks
        worksheet.VerticalPageBreaks.Clear();

        // Save the workbook without any manual page breaks
        workbook.Save("NoPageBreaks.xlsx");
    }
}
