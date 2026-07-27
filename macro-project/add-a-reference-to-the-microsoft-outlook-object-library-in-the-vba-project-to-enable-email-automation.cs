using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaOutlookReference
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (macro-enabled will be saved later)
            Workbook workbook = new Workbook();

            // Access the VBA project (automatically created for .xlsm files)
            VbaProject vbaProject = workbook.VbaProject;

            // Set project encoding (optional, but good practice)
            vbaProject.Encoding = Encoding.UTF8;

            // Add a reference to the Microsoft Outlook Object Library.
            // The libid string below is the typical identifier for Outlook 16.0.
            // Adjust the path/version if a different Outlook version is targeted.
            string outlookName = "Outlook";
            string outlookLibId = "*\\G{00062FFF-0000-0000-C000-000000000046}#9.0#0#C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB#Microsoft Outlook 16.0 Object Library";

            vbaProject.References.AddRegisteredReference(outlookName, outlookLibId);

            // Optionally, add a simple VBA module that uses Outlook (demonstration purpose)
            int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Class, "OutlookHelper");
            VbaModule module = vbaProject.Modules[moduleIndex];
            module.Codes = @"
Public Sub SendTestMail()
    Dim outlookApp As Object
    Set outlookApp = CreateObject(""Outlook.Application"")
    Dim mail As Object
    Set mail = outlookApp.CreateItem(0) ' 0 = olMailItem
    mail.Subject = ""Test Email""
    mail.Body = ""This is a test email sent from VBA.""
    mail.To = ""example@example.com""
    mail.Send
End Sub
";

            // Save the workbook as a macro-enabled file
            workbook.Save("WorkbookWithOutlookReference.xlsm", SaveFormat.Xlsm);

            Console.WriteLine("Workbook saved with Outlook reference added.");
        }
    }
}