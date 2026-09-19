// Title: Save an Excel workbook with embedded OLE objects using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an existing .xlsx file and saves it with Aspose.Cells while preserving all embedded OLE objects. | Show how to use Aspose.Cells SaveFormat.Xlsx to export a workbook without losing OLE object properties. | Provide a snippet that demonstrates persisting inserted OLE objects by saving the workbook to a new file in C#.
// Common Searches: Aspose.Cells C# save workbook without stripping OLE objects | How to keep embedded OLE object properties when exporting Excel with Aspose.Cells | C# Aspose.Cells SaveFormat.Xlsx preserving OLE object | Load and re‑save Excel file containing OLE object using Aspose.Cells .NET | Export Excel workbook with OLE objects using Aspose.Cells API
// Tags: Aspose.Cells save workbook with OLE objects | C# preserve embedded OLE properties | SaveFormat.Xlsx OLE object retention | Excel OLE object export .NET | Aspose.Cells embedded OLE handling

using Aspose.Cells;

// The example loads an existing Excel workbook that already contains an OLE object and saves it as a new XLSX file using Aspose.Cells for .NET, ensuring that the OLE object and its properties are retained.
class Program
{
    static void Main()
    {
        // Load the existing workbook that already contains the OLE object
        Workbook workbook = new Workbook("InputWorkbook.xlsx");

        // Save the workbook to persist the OLE object and its properties
        workbook.Save("OutputWorkbook.xlsx", SaveFormat.Xlsx);
    }
}
