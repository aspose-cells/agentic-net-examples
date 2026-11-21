// Title: Check if a VBA project in an Excel workbook is digitally signed using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# how to read VBA project IsSigned property | determine whether an Excel macro is digitally signed with Aspose.Cells | C# code sample to verify VBA project signature status in .xlsx | using Aspose.Cells to check digital signature of VBA project in workbook | read VBA digital signature flag from workbook with Aspose.Cells .NET
// Tags: Aspose.Cells VBA signature flag | C# check VBA signed status | Excel workbook macro signing detection | retrieve VBA project properties Aspose.Cells | macro signature verification .NET

using System;
using Aspose.Cells;

// The example loads an Excel workbook, accesses the IsSigned property of its VBA project via Aspose.Cells, and prints whether the VBA project is digitally signed.
class Program
{
    static void Main()
    {
        // Load the workbook that may contain a VBA project
        Workbook workbook = new Workbook("input.xlsx");

        // Check if the VBA project is digitally signed
        bool isSigned = workbook.VbaProject.IsSigned;

        // Output the result
        Console.WriteLine($"VBA project signed: {isSigned}");
    }
}
