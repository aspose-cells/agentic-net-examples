// Title: Set the default Latin (Western) font to Calibri while keeping the FarEast font unchanged in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# change workbook default Latin font to Calibri without affecting FarEast font | how to keep FarEast font unchanged when updating western font in Excel with Aspose.Cells | set default western font Calibri in an existing .xlsx using Aspose.Cells .NET API | modify default style font for Latin characters only in Aspose.Cells workbook | C# example for preserving FarEast font while changing default Latin font in Excel file
// Tags: Aspose.Cells default style Latin font | preserve FarEast font Aspose.Cells | set western font Calibri C# | modify workbook default font .NET | Excel default style font customization

using Aspose.Cells;
using System;

// The sample loads an existing Excel workbook, accesses its default style, sets the Latin (Western) font name to "Calibri" while leaving the FarEast font untouched, reassigns the updated style to the workbook, and saves the changes.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Retrieve the default style (used for western text)
        Style defaultStyle = workbook.DefaultStyle;

        // Set the Latin (western) font name to Calibri.
        // The FarEast font name is left unchanged.
        defaultStyle.Font.Name = "Calibri";

        // Apply the modified style back to the workbook
        workbook.DefaultStyle = defaultStyle;

        // Save the updated workbook
        workbook.Save("output.xlsx");
    }
}
