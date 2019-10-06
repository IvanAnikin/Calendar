using System;

using System.Reflection;     // to use Missing.Value

//TO DO: If you use the Microsoft Outlook 11.0 Object Library, uncomment the following line.
using Outlook = Microsoft.Office.Interop.Outlook;

using OpenQA.Selenium;

using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

namespace Calendar
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello, sir");


            string homeURL = "https://bakalari.gymjh.cz/next/rozvrh.aspx?s=next";
            IWebDriver driver = new ChromeDriver("Resources");

            driver.Navigate().GoToUrl(homeURL);
            WebDriverWait wait = new WebDriverWait(driver, System.TimeSpan.FromSeconds(5));

        }


        public void addAppointment()
        {
            try
            {
                Outlook.Application oApp = new Outlook.Application();

                // Get the NameSpace and Logon information.
                // Outlook.NameSpace oNS = (Outlook.NameSpace)oApp.GetNamespace("mapi");
                Outlook.NameSpace oNS = oApp.GetNamespace("mapi");

                //Log on by using a dialog box to choose the profile.
                oNS.Logon(Missing.Value, Missing.Value, true, true);

                //Alternate logon method that uses a specific profile.
                // TODO: If you use this logon method, 
                // change the profile name to an appropriate value.
                oNS.Logon("ivan.anikin@outlook.com", "AzaZ135619009", false, true);

                // Get the Calendar folder.
                Outlook.MAPIFolder oCalendar = oNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);

                Outlook.Items oItems = oCalendar.Items;
                Outlook.AppointmentItem appItem = oItems.Add(Outlook.OlItemType.olAppointmentItem) as Outlook.AppointmentItem;
                appItem.AllDayEvent = false;
                appItem.Start = new DateTime(2019, 10, 5, 9, 30, 0);
                appItem.End = new DateTime(2019, 10, 5, 11, 0, 0);
                appItem.Subject = "Coding";
                appItem.Body = "Calendar events";
                appItem.Save();
                appItem.Display(true);
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public void printFistAppointmentInfo()
        {
            try
            {
                Outlook.Application oApp = new Outlook.Application();

                // Get the NameSpace and Logon information.
                // Outlook.NameSpace oNS = (Outlook.NameSpace)oApp.GetNamespace("mapi");
                Outlook.NameSpace oNS = oApp.GetNamespace("mapi");

                //Log on by using a dialog box to choose the profile.
                oNS.Logon(Missing.Value, Missing.Value, true, true);

                //Alternate logon method that uses a specific profile.
                // TODO: If you use this logon method, 
                // change the profile name to an appropriate value.
                oNS.Logon("ivan.anikin@outlook.com", "AzaZ135619009", false, true);

                // Get the Calendar folder.
                Outlook.MAPIFolder oCalendar = oNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);

                // Get the Items (Appointments) collection from the Calendar folder.
                Outlook.Items oItems = oCalendar.Items;

                // Get the first item.
                Outlook.AppointmentItem oAppt = (Outlook.AppointmentItem)oItems.GetFirst();


                Console.WriteLine("Subject: " + oAppt.Subject);
                Console.WriteLine("Organizer: " + oAppt.Organizer);
                Console.WriteLine("Start: " + oAppt.Start.ToString());
                Console.WriteLine("End: " + oAppt.End.ToString());
                Console.WriteLine("Location: " + oAppt.Location);
                Console.WriteLine("Recurring: " + oAppt.IsRecurring);


                oAppt.Display(true);

                // Done. Log off.
                oNS.Logoff();

                // Clean up.
                oAppt = null;
                oItems = null;
                oCalendar = null;
                oNS = null;
                oApp = null;
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
            }
        }

    }
}
