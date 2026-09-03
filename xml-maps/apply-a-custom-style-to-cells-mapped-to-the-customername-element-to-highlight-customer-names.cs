// Title: Apply a custom yellow background and red bold font style to the 'Customer_Name' named range (mapped to /Customer/Name) using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a Style with a solid yellow fill, red bold font, and applies it to the named range 'Customer_Name' in a workbook loaded with Aspose.Cells. | Write a C# snippet that loads an existing Excel template, defines a custom highlight style, uses StyleFlag.All, applies the style to the XML‑mapped range for Customer Name, and saves the workbook.
// Common Searches: Aspose.Cells C# apply custom style to named range linked to XML element | highlight cells mapped to /Customer/Name in Excel using Aspose.Cells | C# set yellow background and red bold font for Customer_Name range with Aspose.Cells | how to use StyleFlag.All when styling XML mapped cells in Aspose.Cells .NET | apply formatting to XML mapped range Aspose.Cells C#
// Tags: apply custom style to named range Aspose.Cells | highlight XML mapped cells C# | StyleFlag.All usage Aspose.Cells | yellow background red font style Aspose.Cells | Customer_Name range styling Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The program loads a template workbook (or creates a new one), creates a Style with a solid yellow fill and red bold font, sets a StyleFlag with All=true, applies the style to the named range "Customer_Name" (mapped to the /Customer/Name XML element), and saves the modified workbook as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "Template.xlsx";
            const string outputPath = "Output.xlsx";

            // Load existing template or create a new workbook if the file is missing
            Workbook workbook = File.Exists(templatePath) ? new Workbook(templatePath) : new Workbook();

            // Create a custom style to highlight customer names
            Style highlightStyle = workbook.CreateStyle();
            highlightStyle.ForegroundColor = Color.Yellow;
            highlightStyle.Pattern = BackgroundType.Solid;
            highlightStyle.Font.Color = Color.Red;
            highlightStyle.Font.IsBold = true;

            // Apply all style attributes
            StyleFlag flag = new StyleFlag { All = true };

            // Apply the style to the named range "Customer_Name"
            Worksheet sheet = workbook.Worksheets[0];
            Aspose.Cells.Range nameRange = sheet.Cells.CreateRange("Customer_Name");
            nameRange.ApplyStyle(highlightStyle, flag);

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
