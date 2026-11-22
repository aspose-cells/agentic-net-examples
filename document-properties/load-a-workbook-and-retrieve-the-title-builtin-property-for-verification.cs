// Title: Read the built‑in Title property from an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: how to get the Title built‑in document property from an .xlsx file with Aspose.Cells in C# | Aspose.Cells C# example for reading workbook built‑in Title metadata | retrieve Excel file Title using Aspose.Cells .NET API | C# code to access built‑in document properties of an Excel workbook with Aspose.Cells | sample code for reading the Title property of a workbook via Aspose.Cells
// Tags: read built‑in Title property Aspose.Cells | access Excel document metadata C# | Aspose.Cells workbook built‑in properties | retrieve Excel file Title using .NET | load workbook and get metadata Aspose.Cells

using Aspose.Cells;
using System;

// The example loads an Excel workbook from 'input.xlsx' with Aspose.Cells, accesses the built‑in Title property via workbook.BuiltInDocumentProperties.Title, and prints the title to the console for verification.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Retrieve the built‑in Title property
        string title = workbook.BuiltInDocumentProperties.Title;

        // Output the title for verification
        Console.WriteLine("Workbook Title: " + title);
    }
}
