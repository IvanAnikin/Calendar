using System;

using System.Reflection;     // to use Missing.Value

//TO DO: If you use the Microsoft Outlook 11.0 Object Library, uncomment the following line.
using Outlook = Microsoft.Office.Interop.Outlook;

using OpenQA.Selenium;

using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System.Collections.Generic;

namespace ConsoleApp1
{

    public class Subject
    {
        public static string subject { get; set; }
        public static DateTime dateTimeStart;
        public static string notes;

        //public string getSubject() { return subject; }
        //public void putSubject(string subjectName) { subjectName = subject; }
    }

    public class Day
    {
        public static Subject[] subjects;
        public static DateTime dateTimeStart;
        public static string notes;
    }

    public class Week
    {
        public static Day[] days;
        public DateTime dateTimeStart { get; set; }
        public static string notes;

    }

    class Program
    {
        public static string logInUrl = "https://bakalari.gymjh.cz/next/login.aspx";

        public static IWebDriver driver = new ChromeDriver();

        public static void addAppointment()
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

        public static void printFistAppointmentInfo()
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

        public static void logIn(Boolean rememberMe)
        {
            

            driver.Navigate().GoToUrl(logInUrl);

            driver.FindElement(By.Id("username")).SendKeys("ANIKI91843");
            driver.FindElement(By.Id("password")).SendKeys("1f7nv1qe");

            if (rememberMe) driver.FindElement(By.XPath("//*[@id='labelpersistent']")).Click();

            driver.FindElement(By.Id("loginButton")).Click();

        }

        public static void goToKlasifikace()
        {
            driver.FindElement(By.XPath("//*[@id='panel']/div/nav/ul/li[6]/a")).Click();
            driver.FindElement(By.XPath("//*[@id='panel']/div/nav/ul/li[6]/ul/li[1]/a")).Click();
        }

        public static void goToRozvrh()
        {
            driver.FindElement(By.XPath("//*[@id='panel']/div/nav/ul/li[8]/a")).Click();
            driver.FindElement(By.XPath("//*[@id='panel']/div/nav/ul/li[8]/ul/li[1]/a")).Click();
        }

        public static void navigateToRozvrh(Boolean nextWeek)
        {
            if(nextWeek) driver.Navigate().GoToUrl("https://bakalari.gymjh.cz/next/rozvrh.aspx?s=next");
            else driver.Navigate().GoToUrl("https://bakalari.gymjh.cz/next/rozvrh.aspx");
        }

        public static Week getWeek()
        {
            Week week = new Week();
            DateTime today = DateTime.Now;

            IList<IWebElement> weekList = driver.FindElements(By.ClassName("day-row"));

            int i = 0;
            int date = 0;
            int month = 0;
            string firstDate = "";
            foreach(IWebElement element in weekList)
            {
                //Console.WriteLine(element.FindElement(By.ClassName("day-name")).FindElement(By.TagName("span")).Text);
                if(i == 0)
                {
                    week.dateTimeStart = DateTime.Now;
                    firstDate = element.FindElement(By.ClassName("day-name")).FindElement(By.TagName("span")).Text;
                    if (firstDate.ToCharArray()[1] == '.')
                    {
                        date = int.Parse(firstDate.Substring(0, 1));
                        if (firstDate.ToCharArray()[3] == '.') month = int.Parse(firstDate.Substring(2, 1));
                        else month = int.Parse(firstDate.Substring(2, 2));
                    }
                    else
                    {
                        date = int.Parse(firstDate.Substring(0, 2));
                        if (firstDate.ToCharArray()[4] == '.') month = int.Parse(firstDate.Substring(3, 1));
                        else month = int.Parse(firstDate.Substring(3, 2));
                    }
                    week.dateTimeStart = new DateTime(today.Year, month, date, 0, 0, 0);
                }
                i++;
            }

            // XPATH OF FIRST DATE --- GET ALL BY XPATH ??? - OR LOOPS??? //*[@id="schedule"]/div/div[2]/div/div/div/div/span

            return week;
        }

        public static Day getDay()
        {
            IList<IWebElement> subjectsList = driver.FindElements(By.ClassName(""));

            Day day = new Day();

            /*
            int i = 0;
            foreach (IWebElement subjectElement in subjectsList)
            {
                day.subjects[i].putSubject("");
                i++;
            }
            */
            return day;
        }

        public static Subject getSubject()
        {
            Subject subject = new Subject();



            return subject;
        }

        static void Main(string[] args)
        {
            Console.WriteLine("Hello, sir");

            logIn(true);

            navigateToRozvrh(false);

            Console.WriteLine(getWeek().dateTimeStart);

        }

        

    }
}
