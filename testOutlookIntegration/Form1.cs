using System;
using System.Collections.Generic;
using System.Globalization;
using System.ComponentModel;
using System.Data;
using System.Net;
using System.IO;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.XPath;
using Microsoft.Office.Interop.Outlook;
using HtmlAgilityPack;

namespace testOutlookIntegration
{
    public partial class Form1 : Form
    {
        private Microsoft.Office.Interop.Outlook.Application OLapp;

        public Form1()
        {
            InitializeComponent();

            log("Opening Outlook");
            OLapp = new Microsoft.Office.Interop.Outlook.Application();

            log("Initialising Application");
            loadWebComponent();     
            
            log(this.Text + " is ready");
        }

        private void lockForm()
        {
            weekList.Enabled = false;
            
            this.Cursor = Cursors.WaitCursor;
        }

        private void unlockForm()
        {
            weekList.Enabled = true;
            this.Cursor = Cursors.Default;
        }

        private void log(string msg)
        {
            if (!msg.EndsWith("\n"))
                msg += "\r\n";
            logBox.Text += msg;
            Update();
        }

        private void doSync(String url)
        {
            lockForm();

            // first, remove any old apps that were made previously by this application:
            log("Removing old Outlook appointments created with this application...");
            removeOldAppointments();

            Regex rgxWeeks = new Regex("weeks=\\d+");
            Regex rgxDays = new Regex("days=\\d+\\-\\d+");
            Regex rgxPeriods = new Regex("periods=\\d+\\-\\d+");

            Calendar cal = CultureInfo.InvariantCulture.Calendar;
            int weekNow = cal.GetWeekOfYear(DateTime.Now, System.Globalization.CalendarWeekRule.FirstDay,DayOfWeek.Monday);
            for (int i = 0; i < weekList.Value; i++)
            {
                log("Synchronising appointments for week " + (weekNow + i));
                // find part of URL that specifies the week, and loop around for all weeks between now and now + i weeks:
                String weekURL = rgxWeeks.Replace(url, "weeks=" + (weekNow + i));
                weekURL = rgxDays.Replace(weekURL, "days=1-7");
                weekURL = rgxPeriods.Replace(weekURL, "periods=1-48");
                scrapeSyllabusPlus(weekURL, weekNow + i);
            }

            log("\r\nAll Appointments have been Synchronised!");

            unlockForm();
        }

        private void removeOldAppointments()
        {
            NameSpace ns = OLapp.GetNamespace("MAPI");
            //MAPIFolder f = ns.PickFolder();
            MAPIFolder rootFolder = ns.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderCalendar);
            for (int i = rootFolder.Items.Count; i > 0; i--)
            //foreach (AppointmentItem item in items)
            {
                AppointmentItem item = (AppointmentItem) rootFolder.Items[i];
                if ( item.Body != null && item.Body != "" && item.Body.Contains("Created by " + this.Text))
                {
                    item.Delete();
                    // remove item:
                }
            }
        }

        private void scrapeSyllabusPlus(String url, int weekNumber)
        {
            String webPage = getWebpage(url);
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(webPage);

            // get an XPath navigator so we can browse the document:
            XPathNavigator nav = doc.CreateNavigator();
            // the second table is the one we care about:
            XPathNodeIterator it = nav.Select("/html/body/table[2]");
            it.MoveNext();

            XPathNavigator rootTbl = it.Current;
            XPathNodeIterator rows = rootTbl.Select("tr");
            bool firstRow = true;
            List<String> times = new List<String>();
            String day = "";

            foreach (XPathNavigator row in rows)
            {
                XPathNodeIterator cols = row.Select("td");

                int currentCell = 0;
                bool firstCol = true;
                foreach (XPathNavigator col in cols)
                {
                    // first row contains times. It would be nice if the first row was tagged with a CSS ID, but there you go..
                    if (firstRow)
                    {
                        times.Add(col.ToString());
                    }
                    else
                    {
                        // if the current cell has CSS class "row-label-one" then skip it - it's the day of week header
                        // although we may want to keep this so we know the date? nah...
                        if (firstCol)
                        {
                            firstCol = false;
                            day = col.ToString();
                            ++currentCell;
                            continue;
                        }
                        // if the current cell has CSS class "object-cell-border then this is an appointment that needs to be
                        // synced! 
                        if (col.Select("table").Count > 0)
                        //if (col.GetAttribute("class", "") == "object-cell-border")
                        {
                            // this is an event we need to sync:
                            // start time is the current cell lication:
                            String startTime = times.ElementAt(currentCell);
                            // end time is the current cell location plus colspan attribute:
                            int colspan = Int32.Parse(col.GetAttribute("colspan", ""));
                            String endTime = times.ElementAt(currentCell + colspan);

                            // there are three embedded <table> elements.
                            // the first one has the generic subject, like "Bachelor of Information Technology".
                            String department = getStringFromXSLTPath(col, "table[1]/tr/td");
                            // the second has the specific subject and type, like "Web Fundamentals", "Lecture"
                            String subject = getStringFromXSLTPath(col, "table[2]/tr/td[1]");
                            String subjType = getStringFromXSLTPath(col, "table[2]/tr/td[2]");
                            // the third has the weeks and room info.
                            String room = getStringFromXSLTPath(col, "table[3]/tr/td[2]");

                            // work out the date we're on. We know the week we're in, and we can get the week day number easily enough...
                            DateTime startDT = getDateTimeFromDayAndWeek(day, weekNumber, startTime);
                            DateTime endDT = getDateTimeFromDayAndWeek(day, weekNumber, endTime);

                            createCalendarEvent(startDT, endDT, department, subject, subjType, room);

                            // finished processing, so add the current colspan to the current cell number:
                            currentCell += colspan;
                        }
                        else
                        {
                            ++currentCell;
                        }
                    }
                }
                // completed at least one row:
                firstRow = false;
            }
        }

        private void createCalendarEvent(DateTime start, DateTime end, String dept, String subj, String subjType, String room)
        {
            // TODO - make this configurable!
            if (end < DateTime.Now)
                return;

            AppointmentItem apt = (AppointmentItem)OLapp.CreateItem(OlItemType.olAppointmentItem);
            
            apt.Start = start;
            apt.End = end;
            apt.Subject = subj + " - " + subjType;
            apt.Body = "Subject: " + subj + " (" + subjType + ")"
                    + "\nDepartment: " + dept + "\nRoom: " + room
                    + "\n\nCreated by " + this.Text;
            apt.Location = room;
            //MAPIFolder folder = getSelectedFolder();
            apt.Save();
            //folder.Items.Add(apt);
            
            
            
        }

        private DateTime getDateTimeFromDayAndWeek(String dayName, int weekNumber, String timeString)
        {
            Calendar myCal = CultureInfo.InvariantCulture.Calendar;
            DateTime startOfYear = new DateTime(DateTime.Now.Year, 1, 1);
            DateTime dt = myCal.AddWeeks(startOfYear, weekNumber - 1);

            #region dow
            DayOfWeek dow = new DayOfWeek();
            switch (dayName.ToUpper())
            {
                case "MON":
                    dow = DayOfWeek.Monday;
                    break;
                case "TUE":
                    dow = DayOfWeek.Tuesday;
                    break;
                case "WED":
                    dow = DayOfWeek.Wednesday;
                    break;
                case "THU":
                    dow = DayOfWeek.Thursday;
                    break;
                case "FRI":
                    dow = DayOfWeek.Friday;
                    break;
                case "SAT":
                    dow = DayOfWeek.Saturday;
                    break;
                case "SUN":
                    dow = DayOfWeek.Sunday;
                    break;
            }
            #endregion 

            dt = dt.AddDays(dow - dt.DayOfWeek);
            dt += TimeSpan.Parse(timeString);

            return dt;
        }

        private String getStringFromXSLTPath(XPathNavigator node, String XSLTSelector)
        {
            XPathNodeIterator it = node.Select(XSLTSelector);
            it.MoveNext();
            return it.Current.ToString();
        }

        private String getWebpage(String uri)
        {
            StringBuilder sb = new StringBuilder();
            // used on each read operation
            byte[] buf = new byte[8192];

            // prepare the web page we will be asking for
            WebRequest request = WebRequest.Create(uri);

            // execute the request
            WebResponse response = request.GetResponse();

            // we will read data via the response stream
            Stream resStream = response.GetResponseStream();

            string tempString = null;
            int count = 0;

            do
            {
                // fill the buffer with data
                count = resStream.Read(buf, 0, buf.Length);

                // make sure we read some data
                if (count != 0)
                {
                    // translate from bytes to ASCII text
                    tempString = Encoding.ASCII.GetString(buf, 0, count);

                    // continue building the string
                    sb.Append(tempString);
                }
            }
            while (count > 0); // any more data to read?

            // print out page source
            return sb.ToString();
        }

        private void loadWebComponent()
        {
            // need to grab the SPLus web page, remove the bits we don't want, and load it into 
            // our web component. 
            //
            // We want to remove all the date controls, and replace the JS that loads the timetable into
            // a new page. The new JS will pass the URL back to us here. Yay!


            // load web page:
            HtmlAgilityPack.HtmlWeb webLoader = new HtmlAgilityPack.HtmlWeb();
            HtmlAgilityPack.HtmlDocument doc = webLoader.Load("http://webit.tekotago.ac.nz/SPlus/TT/Splus.cfm?TTvalue=staff");

            XPathNodeIterator it;

            #region JS_REPLACE

            // insert our overloaded JS stuff:
            String newJS = @"
function viewTimetable2()
{
 
	if (document.form.identifier.value=="" "")
		alert(""Please select a staff members name."");
	else
	{				
		var location;
		location = ""http://splus.op.ac.nz:8080/reporting/individual?&identifier=""+document.form.identifier.value+""&idtype=id&objectclass=staff&width=100""+CheckWeeks();
        window.external.loadTimetable(location);
	}
}
";

            // create new script node:
            HtmlNode n = doc.CreateElement("script");
            n.SetAttributeValue("type", "text/javascript");
            n.InnerHtml = newJS;

            // insert new script node right after first div node:
            HtmlNode bodyTag = doc.DocumentNode.SelectSingleNode("/html");
            HtmlNode divTag = doc.DocumentNode.SelectSingleNode("/html/div");
            bodyTag.InsertBefore(n, divTag);

            #endregion

            // remove unwanted table rows:
            HtmlNode parentNode = doc.DocumentNode.SelectSingleNode("/html/div/table[3]");
            foreach (HtmlNode cn in parentNode.SelectNodes("tr[position()<6]"))
            //foreach (HtmlNode cn in parentNode.SelectNodes("tr"))
            {
                if (cn != null)
                    parentNode.RemoveChild(cn, false);
            }

            // remove unwanted images:
            foreach (HtmlNode cn in doc.DocumentNode.SelectNodes("//img"))
            {
                if (cn != null)
                {
                    HtmlNode parent = cn.SelectSingleNode("..");
                    if (parent != null)
                    {
                        parent.RemoveChild(cn);
                    }
                }
            }

            // remove link to parent page:
            foreach (HtmlNode cn in doc.DocumentNode.SelectNodes("//a[@href='view.cfm']"))
            {
                if (cn != null)
                {
                    HtmlNode parent = cn.SelectSingleNode("..");
                    if (parent != null)
                    {
                        parent.RemoveChild(cn);
                    }
                }
            }
            

            //  change input button:
            HtmlNode inputNode = doc.DocumentNode.SelectSingleNode("//input[@name='goto']");
            inputNode.SetAttributeValue("value", "Synchronise Timetable");
            //inputNode.SetAttributeValue("onClick", "viewTimetable2()");
            inputNode.SetAttributeValue("onClick", "viewTimetable2()");

            StringWriter sw = new StringWriter();
            doc.Save(sw);
            web.DocumentText = sw.ToString();

            web.ObjectForScripting = this;
        }

        public void loadTimetable(String url)
        {
            doSync(url);
        }
        private void web_Navigating(object sender, WebBrowserNavigatingEventArgs e)
        {
            // picking last name letter counts as navigation. Allow it:
            if (e.Url.AbsoluteUri.StartsWith("javascript"))
                e.Cancel = false;
            else
                // but disallow anything else
                e.Cancel = true;
         
        }

    }
}
