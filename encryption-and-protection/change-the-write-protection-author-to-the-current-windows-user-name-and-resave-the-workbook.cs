// Title: How to set the write‑protection author to the current Windows user and save the workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells change workbook write protection author to Environment.UserName | Set Excel file write‑protection author to logged‑in Windows user with Aspose.Cells | Save an Excel workbook after updating write protection metadata in .NET | Aspose.Cells update write protection author programmatically in C# | How to preserve write‑protection settings when re‑saving an Excel file with Aspose.Cells
// Tags: Aspose.Cells write protection author | C# set workbook write protection metadata | environment username Excel protection | Aspose.Cells save workbook after protection change | Excel write protection author .NET

using System;
using Aspose.Cells;

// The example loads an existing Excel workbook with Aspose.Cells, assigns the current Windows user name to the workbook's write‑protection author via workbook.Settings.WriteProtection.Author, and then saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Change the write‑protection author to the current Windows user name
        workbook.Settings.WriteProtection.Author = Environment.UserName;

        // Re‑save the workbook (you can overwrite or specify a new file name)
        workbook.Save("output.xlsx");
    }
}
