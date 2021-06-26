using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace Aetna_Scraper
{

    public class Dentist
    {
        public string CitySection { get; set; }
        public string Name { get; set; }
        public string Address1 { get; set; }
        public string Address2 { get; set; }
        public string City { get; set; }
        public string St { get; set; }
        public string Zip { get; set; }
        public string Phone { get; set; }
        public string Fax { get; set; }
        public string Other { get; set; }
        public string Specialty { get; set; }
        public string URL { get; set; }

    }

    public class Program
    {
        // Start Chrome windows (main and secondary)
        IWebDriver mainbrowser = new ChromeDriver(@"C:\Program Files\chromedriver");
        IWebDriver popwindow = new ChromeDriver(@"C:\Program Files\chromedriver");

        // Excel Objects
        Microsoft.Office.Interop.Excel.Application oXL;
        Microsoft.Office.Interop.Excel._Workbook oWB;
        Microsoft.Office.Interop.Excel._Worksheet oSheet;

        public int excelRow = 2;
        List<string> cities = new List<string>();
        //static void Main(string[] args)
        //{
        //    Program prog = new Program();
        //    //prog.StateSelect("http://www.nationaldirectoryofdentists.com/dentists/");
        //    prog.DentistParser();

        //}

        public void TestMethod()
        {
            popwindow.Navigate().GoToUrl("https://www.whitepages.com/");

        }

        public void DentistParser()
        {

            // Set a variable to the Documents path.
            string docPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            // Write the string array to a new file named "WriteLines.txt".
            //StreamWriter outputFile = new StreamWriter(Path.Combine(docPath, "2020 AJGA Events.CSV"));

            var outputFile = new FileInfo(Path.Combine(docPath, "AetnaDentists.xlsx"));
            if (File.Exists(outputFile.ToString()))
            {
                File.Delete(outputFile.ToString());
            }


            object misvalue = System.Reflection.Missing.Value;

            // Start Excel and get Application object.
            oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = true;

            //Get a new workbook.
            oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
            oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

            mainbrowser.Manage().Window.Size = new Size(1300, 1500);
            popwindow.Manage().Window.Size = new Size(1300, 1500);

            List<Dentist> PlayerList = new List<Dentist>();

            oXL.Cells[1, 1] = "CityURL";
            oXL.Cells[1, 2] = "Name";
            oXL.Cells[1, 3] = "Address";
            oXL.Cells[1, 4] = "City";
            oXL.Cells[1, 5] = "St";
            oXL.Cells[1, 6] = "Zip";
            oXL.Cells[1, 7] = "Phone";
            oXL.Cells[1, 8] = "Fax";
            oXL.Cells[1, 9] = "Other";
            oXL.Cells[1, 10] = "Specialty";
            oXL.Cells[1, 11] = "URL";

            // Parse City
            //StateSelect("https://www.aetna.com/dsepublic/#/contentPage?page=providerResults&parameters=searchText%3D''All%20Dental%20Professionals'';isGuidedSearch%3Dtrue&site_id=DirectLinkDental&language=en");
            for (int i = 0; i < cities.Count; i++)
            {
                CityParse(cities[i]);
            }

            //CityParse("http://www.nationaldirectoryofdentists.com/dentists/OH/Canton?page=9");
            //CityParse("http://www.nationaldirectoryofdentists.com/dentists/IN/Muncie");

            oWB.SaveAs(outputFile, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            oWB.Close();
            oXL.Quit();

            if (mainbrowser != null) { mainbrowser.Dispose(); }
            if (popwindow != null) { popwindow.Dispose(); }

        }
        public void StateSelect()
        {



        }
        public void CityParse(string URL)
        {
            mainbrowser.Navigate().GoToUrl(URL);

            // Check for Next Button
            int currentPage = 1;
            Boolean NextPageExists = true;
            // Loop through additional pages if necessary

            while (NextPageExists)
            {
                // Get Jobs
                Console.WriteLine(string.Format("--------------------- Parsing Page {0} ---------------------", currentPage.ToString()));

                // parse page
                string JobRowXPath = "//div[@class='dataGridTableContant']/li[@class='col-xs-12 pad0 dataGridRow customPurpleRecord']}";
                foreach (IWebElement jobrow in mainbrowser.FindElements(By.XPath(JobRowXPath)))
                {
                    // Pull job row for data grab
                    Dentist drec = new Dentist();

                    string CityStZip = "";
                    // Rip Data
                    try { drec.CitySection = URL; } catch { }
                    try { drec.Name = mainbrowser.FindElement(By.XPath("//a[@class='providerNameDetails']")).Text; } catch { }
                    try
                    {
                        CityStZip = mainbrowser.FindElement(By.XPath("//div[@ng-if='provider.providerLocations.address.streetLine2']")).Text.Replace("\r", "");
                    }
                    catch { }
                    try { drec.Address1 = mainbrowser.FindElement(By.XPath("//span[@ng-bind-html='provider.providerLocations.address.streetLine1|trustHtml']")).Text; } catch { }
                    try { drec.Address2 = mainbrowser.FindElement(By.XPath("//span[@ng-bind-html='provider.providerLocations.address.streetLine2|trustHtml']")).Text; } catch { }
                    try { drec.City = CityStZip.Split(',')[0]; } catch { }
                    try { drec.St = CityStZip.Split(',')[1].Trim().Split(' ')[0]; } catch { }
                    try { drec.Zip = CityStZip.Split(',')[1].Trim().Split(' ')[1]; } catch { }
                    try { drec.Phone = mainbrowser.FindElement(By.XPath("//span[@class='dataGridPadding padL3']")).Text; } catch { }
                    try { drec.Specialty = mainbrowser.FindElement(By.XPath("//span[@ng-bind-html='spec.specialty.description | trustHtml']")).Text; } catch { }

                    // Place in Excel Row
                    oXL.Cells[excelRow, 1] = drec.CitySection;
                    oXL.Cells[excelRow, 2] = drec.Name;
                    oXL.Cells[excelRow, 3] = drec.Address1;

                    oXL.Cells[excelRow, 4] = drec.City;
                    oXL.Cells[excelRow, 5] = drec.St;
                    oXL.Cells[excelRow, 6] = drec.Zip;
                    oXL.Cells[excelRow, 7] = drec.Phone;
                    oXL.Cells[excelRow, 8] = drec.Fax;
                    oXL.Cells[excelRow, 9] = drec.Other;
                    oXL.Cells[excelRow, 10] = drec.Specialty;

                    excelRow += 1;
                }

                try
                {
                    // Check for Next Button
                    NextPageExists = mainbrowser.FindElements(By.XPath("//a[@class='next-link']")).Count > 0;

                    // Click Next page button
                    if (NextPageExists)
                    {
                        mainbrowser.FindElement(By.XPath("//a[@class='next-link']")).Click();
                    }

                    currentPage++;
                    URL = mainbrowser.Url;

                }
                catch { }


            }

        }

    }
}
