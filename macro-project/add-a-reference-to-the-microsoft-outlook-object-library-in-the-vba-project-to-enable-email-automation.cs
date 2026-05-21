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
            // Create a new workbook (macro-enabled format will be used when saving)
            Workbook workbook = new Workbook();

            // Access the VBA project (automatically created for .xlsm files)
            VbaProject vbaProject = workbook.VbaProject;

            // Optional: set project name and encoding
            vbaProject.Name = "OutlookAutomationProject";
            vbaProject.Encoding = Encoding.UTF8;

            // Add a reference to the Microsoft Outlook Object Library
            // The libid string below is an example for Outlook 16.0; adjust the path/version as needed.
            string outlookLibId = "*\\G{00062FFF-0000-0000-C000-000000000046}#9.0#0#C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB#Microsoft Outlook 16.0 Object Library";
            vbaProject.References.AddRegisteredReference("Outlook", outlookLibId);

            // (Optional) Add a simple VBA module that uses Outlook – just for demonstration
            int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Class, "EmailHelper");
            vbaProject.Modules[moduleIndex].Codes =
                "Sub SendMail()\r\n" +
                "    Dim olApp As Outlook.Application\r\n" +
                "    Dim olMail As Outlook.MailItem\r\n" +
                "    Set olApp = New Outlook.Application\r\n" +
                "    Set olMail = olApp.CreateItem(0)\r\n" +
                "    olMail.Subject = \"Test Email\"\r\n" +
                "    olMail.Body = \"This email was sent from VBA in an Excel workbook.\"\r\n" +
                "    olMail.Recipients.Add \"example@example.com\"\r\n" +
                "    olMail.Send\r\n" +
                "End Sub";

            // Save the workbook as a macro-enabled file
            workbook.Save("WorkbookWithOutlookReference.xlsm", SaveFormat.Xlsm);
        }
    }
}