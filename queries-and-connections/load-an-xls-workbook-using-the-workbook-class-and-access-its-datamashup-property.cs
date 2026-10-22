// Title: Load an XLS workbook and inspect its DataMashup property using Aspose.Cells for .NET (C#)
// AI Prompts: Load an .xls file with Aspose.Cells Workbook and retrieve the DataMashup object. | Determine whether the workbook contains a DataMashup and print a message.
// Common Searches: how to get DataMashup from an existing .xls file using Aspose.Cells C# | Aspose.Cells Workbook.DataMashup property example in C# | checking if Power Query model exists in loaded workbook Aspose.Cells
// Tags: load xls workbook Aspose.Cells | access Workbook.DataMashup C# | check DataMashup null Aspose.Cells | retrieve Power Query model from XLS | Aspose.Cells DataMashup usage

using System;
using Aspose.Cells;

// The sample creates a Workbook instance from an existing "input.xls" file, accesses its DataMashup property to obtain the underlying data model (e.g., Power Query), and writes to the console whether the DataMashup object is present or null.
class Program
{
    static void Main()
    {
        // Load an existing XLS workbook from file
        Workbook workbook = new Workbook("input.xls");

        // Access the DataMashup property of the workbook
        // DataMashup provides access to the underlying data model (e.g., Power Query data)
        var dataMashup = workbook.DataMashup;

        // Example: output whether a DataMashup is present
        Console.WriteLine("DataMashup is " + (dataMashup != null ? "available" : "null"));
    }
}
