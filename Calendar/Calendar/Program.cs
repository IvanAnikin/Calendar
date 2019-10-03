using System;

using System.Reflection;     // to use Missing.Value

//TO DO: If you use the Microsoft Outlook 11.0 Object Library, uncomment the following line.
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Calendar
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World! How about to try create an outlook event :-)");



            Outlook.Application oApp = new Outlook.Application();

            // Get the NameSpace and Logon information.
            // Outlook.NameSpace oNS = (Outlook.NameSpace)oApp.GetNamespace("mapi");
            Outlook.NameSpace oNS = oApp.GetNamespace("mapi");

            //Log on by using a dialog box to choose the profile.
            oNS.Logon(Missing.Value, Missing.Value, true, true);

            //Alternate logon method that uses a specific profile.
            // TODO: If you use this logon method, 
            // change the profile name to an appropriate value.
            //oNS.Logon("YourValidProfile", Missing.Value, false, true); 

            // Get the Calendar folder.
            Outlook.MAPIFolder oCalendar = oNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);

            // Get the Items (Appointments) collection from the Calendar folder.
            Outlook.Items oItems = oCalendar.Items;

            // Get the first item.
            Outlook.AppointmentItem oAppt = (Outlook.AppointmentItem)oItems.GetFirst();

        }
    }
}
