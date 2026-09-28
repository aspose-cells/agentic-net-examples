// Title: Copy a VBA UserForm (DesignerStorage) from a template Excel workbook to another workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells reflection to duplicate the DesignerStorage object from a source Workbook into a destination Workbook in C#. | Implement a version‑agnostic method that copies a VBA UserForm between two .xlsx files using the Workbook.DesignerStorage.CopyFrom API. | Write C# code that loads a template workbook, extracts its UserForm storage, and injects it into another workbook while handling a missing DesignerStorage property.
// Common Searches: aspnet copy VBA UserForm from one Excel file to another using Aspose.Cells | how to transfer DesignerStorage between workbooks with Aspose.Cells C# | preserve Excel UserForm layout when cloning a workbook in .NET | reflection based DesignerStorage copy Aspose.Cells version compatibility | copy userform storage from template workbook to target workbook Aspose.Cells
// Tags: DesignerStorage.CopyFrom API Aspose.Cells | transfer UserForm data .xlsx C# | reflection access DesignerStorage Aspose.Cells | clone template workbook UserForm .NET | cross‑version Aspose.Cells workbook copy

using System;
using System.IO;
using Aspose.Cells;

// The example loads a template and a target .xlsx file, verifies their existence, accesses the Workbook.DesignerStorage property via reflection, and invokes its CopyFrom method to transfer the VBA UserForm (DesignerStorage) from the template to the target workbook. It then saves the updated workbook, handling cases where the DesignerStorage property is unavailable in the current Aspose.Cells version.
class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "Template.xlsx";
            const string targetPath = "Target.xlsx";
            const string outputPath = "Target_With_UserForm.xlsx";

            // Verify that the input files exist to avoid FileNotFoundException
            if (!File.Exists(templatePath))
                throw new FileNotFoundException($"Template file not found: {templatePath}");
            if (!File.Exists(targetPath))
                throw new FileNotFoundException($"Target file not found: {targetPath}");

            // Load the workbooks
            Workbook templateWb = new Workbook(templatePath);
            Workbook targetWb = new Workbook(targetPath);

            // Copy DesignerStorage (UserForm) using reflection to stay compatible with different Aspose.Cells versions
            var designerProp = typeof(Workbook).GetProperty("DesignerStorage");
            if (designerProp != null)
            {
                var targetDesigner = designerProp.GetValue(targetWb);
                var sourceDesigner = designerProp.GetValue(templateWb);
                var copyFromMethod = targetDesigner?.GetType().GetMethod("CopyFrom");
                copyFromMethod?.Invoke(targetDesigner, new object[] { sourceDesigner });
            }
            else
            {
                Console.WriteLine("DesignerStorage property is not available in this Aspose.Cells version. Skipping UserForm copy.");
            }

            // Save the result
            targetWb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
