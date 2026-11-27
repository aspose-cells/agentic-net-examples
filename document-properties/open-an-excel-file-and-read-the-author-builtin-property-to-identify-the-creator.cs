// Title: Read the Author built‑in document property from an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: c# Aspose.Cells read author built‑in property from xlsx | how to get creator metadata of an Excel file using Aspose.Cells | retrieve Excel workbook built‑in document properties author in .NET | Aspose.Cells example for accessing author property of a spreadsheet | extract Excel file author with Aspose.Cells library in C#
// Tags: retrieve author metadata Aspose.Cells | Aspose.Cells access workbook built‑in document properties | C# obtain spreadsheet creator metadata | extract Excel author using Aspose.Cells API | Aspose.Cells built‑in properties handling in .NET

using System;
using Aspose.Cells;

// The program loads an Excel workbook with Aspose.Cells, accesses the BuiltInDocumentProperties.Author property to obtain the file's author, and prints the author name to the console.
class Program
{
    static void Main()
    {
        // Load the Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Retrieve the Author built‑in property
        string author = workbook.BuiltInDocumentProperties.Author;

        // Output the author name
        Console.WriteLine("Author: " + author);
    }
}
