// Title: Persist an inserted OLE object and its properties by saving a workbook with Aspose.Cells for .NET (C#)
// Common Searches: how to keep embedded OLE objects when saving an Excel workbook using Aspose.Cells C# | Aspose.Cells C# save workbook without losing OLE object properties | C# code to persist OLE object after modifying Excel file with Aspose.Cells | saving workbook as XLSX while preserving embedded OLE objects in Aspose.Cells
// Tags: save workbook with OLE object Aspose.Cells | preserve embedded OLE object C# | Aspose.Cells SaveFormat.Xlsx example | persist OLE object properties Excel | export Excel containing OLE object Aspose.Cells

using Aspose.Cells;
using System;

// The example loads an existing Excel workbook that already contains an OLE object, then saves it as an XLSX file using Aspose.Cells' Save method, ensuring the OLE object and all its associated properties are retained in the output workbook.
class Program
{
    static void Main()
    {
        // Load the workbook that already contains the OLE object
        Workbook workbook = new Workbook("input.xlsx");

        // Save the workbook; this persists the OLE object and all its properties
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
