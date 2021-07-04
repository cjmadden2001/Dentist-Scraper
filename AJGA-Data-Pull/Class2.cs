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
using System.Text.RegularExpressions;
using System.Collections.Specialized;

namespace Aetna_Scraper
{

    public class State
    {
        public State(string abbr, string name, string na, Boolean cs)
        {            
            Name = name;
            Complete = cs;
            NextAlpha = na;
        }
        public string Abbreviation { get; set; }
        public string Name { get; set; }
        public string NextAlpha { get; set; }
        public Boolean Complete { get; set; }

        public override string ToString()
        {
            return string.Format("{0} - {1}", Abbreviation, Name);
        }
    }

    public class Dentist
    {
        public string CitySection { get; set; }
        public string Name { get; set; }
        public string FName { get; set; }
        public string LName { get; set; }
        public string Address { get; set; }
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
       
    public class AetnaProgram
    {
        // Start Chrome windows (main and secondary)
        IWebDriver mainbrowser = new ChromeDriver(@"C:\Program Files\chromedriver");
        //IWebDriver popwindow = new ChromeDriver(@"C:\Program Files\chromedriver");

        // Excel Objects
        Microsoft.Office.Interop.Excel.Application oXL;
        Microsoft.Office.Interop.Excel._Workbook oWB;
        Microsoft.Office.Interop.Excel._Worksheet oSheet;

        public int excelRow = 2;
        List<string> cities = new List<string>();

        string[] stateList = { "Alabama", "Alaska", "Arizona", "Arkansas", "California", "Colorado", "Connecticut", "Delaware", "Florida" };
        // "Alabama", "Alaska", "Arizona", "Arkansas", "California", "Colorado", "Connecticut","Delaware", "Florida",
        // "Georgia", "Hawaii", "Idaho", "Illinois", "Indiana",  "Iowa", "Kansas", "Kentucky", "Louisiana", "Maine", "Maryland", "Massachusetts",
        // "Michigan", "Minnesota", "Mississippi", "Missouri", "Montana", "Nebraska", "Nevada", "New Hampshire",  "New Jersey", "New Mexico", 
        // "New York", "North Carolina", "North Dakota", "Ohio", "Oklahoma", "Oregon", "Pennsylvania", "Rhode Island", "South Carolina", 
        // "South Dakota", "Tennessee", "Texas",  "Utah", 
        // "Vermont", 
        // "Virginia", "Washington", "West Virginia", "Wisconsin", "Wyoming", "District Of Columbia",   

          public static List<State> ListOfStates = new List<State> {
          //us
          new State("AL", "Alabama", "", true),
          new State("AK", "Alaska", "", true),
          new State("AZ", "Arizona", "T", false),
          new State("AR", "Arkansas", "", false),
          new State("CA", "California", "", false),
          new State("CO", "Colorado", "", false),
          new State("CT", "Connecticut", "", false),
          new State("DE", "Delaware", "", false),
          new State("DC", "District Of Columbia", "", false),
          new State("FL", "Florida", "", false),
          new State("GA", "Georgia", "", false),
          new State("HI", "Hawaii", "", false),
          new State("ID", "Idaho", "", false),
          new State("IL", "Illinois", "", false),
          new State("IN", "Indiana", "", false),
          new State("IA", "Iowa", "", false),
          new State("KS", "Kansas", "", false),
          new State("KY", "Kentucky", "", false),
          new State("LA", "Louisiana", "", false),
          new State("ME", "Maine", "", false),
          new State("MD", "Maryland", "", false),
          new State("MA", "Massachusetts", "", false),
          new State("MI", "Michigan", "", false),
          new State("MN", "Minnesota", "", false),
          new State("MS", "Mississippi", "", false),
          new State("MO", "Missouri", "", false),
          new State("MT", "Montana", "", false),
          new State("NE", "Nebraska", "", false),
          new State("NV", "Nevada", "", false),
          new State("NH", "New Hampshire", "", false),
          new State("NJ", "New Jersey", "", false),
          new State("NM", "New Mexico", "", false),
          new State("NY", "New York", "", false),
          new State("NC", "North Carolina", "", false),
          new State("ND", "North Dakota", "", false),
          new State("OH", "Ohio", "", false),
          new State("OK", "Oklahoma", "", false),
          new State("OR", "Oregon", "", false),
          new State("PA", "Pennsylvania", "", false),
          new State("RI", "Rhode Island", "", false),
          new State("SC", "South Carolina", "", false),
          new State("SD", "South Dakota", "", false),
          new State("TN", "Tennessee", "", false),
          new State("TX", "Texas", "", false),
          new State("UT", "Utah", "", false),
          new State("VT", "Vermont", "", false),
          new State("VA", "Virginia", "", false),
          new State("WA", "Washington", "", false),
          new State("WV", "West Virginia", "", false),
          new State("WI", "Wisconsin", "", false),
          new State("WY", "Wyoming", "", false)
          };


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
            //popwindow.Manage().Window.Size = new Size(1300, 1500);

            List<Dentist> PlayerList = new List<Dentist>();

            oXL.Cells[1, 1] = "CityURL";
            oXL.Cells[1, 2] = "Name";
            oXL.Cells[1, 3] = "First";
            oXL.Cells[1, 4] = "Last";
            oXL.Cells[1, 5] = "Addr1";
            oXL.Cells[1, 6] = "Addr2";
            oXL.Cells[1, 7] = "City";
            oXL.Cells[1, 8] = "St";
            oXL.Cells[1, 9] = "Zip";
            oXL.Cells[1, 10] = "Phone";
            oXL.Cells[1, 11] = "Fax";
            oXL.Cells[1, 12] = "Specialty";
            oXL.Cells[1, 13] = "URL";

            // Parse City
            // https://www.aetna.com/dsepublic/#/contentPage?page=providerResults&parameters=searchText%3D'All%20Dental%20Professionals';isGuidedSearch%3Dtrue&site_id=DirectLinkDental&language=en

            string stateSelecURL = "https://www.aetna.com/dsepublic/#/contentPage?page=providerResults&parameters=searchText%3D'All%20Dental%20Professionals';isGuidedSearch%3Dtrue&site_id=DirectLinkDental&language=en";
            mainbrowser.Navigate().GoToUrl(stateSelecURL);

            // Select arbitrary state to get to the starting list
            waitForElement("//input[@id='zip1']");
            mainbrowser.FindElement(By.XPath("//input[@id='zip1']")).SendKeys("Vermont");
            System.Threading.Thread.Sleep(1 * 1000); // 0.5 seconds
            mainbrowser.FindElement(By.XPath("//input[@id='zip1']")).SendKeys(Keys.Tab);
            System.Threading.Thread.Sleep(1*1000); // 0.5 seconds
            mainbrowser.FindElement(By.XPath("//button[@id='second-step-continue']")).Click(); //search button
            
            waitForElement("//a[@class='skipPlanBottom floatR']");
            mainbrowser.FindElement(By.XPath("//a[@class='skipPlanBottom floatR']")).Click();

            foreach (string state in stateList) { 
                StateSelect();
                //StateParse();
            }


            oWB.SaveAs(outputFile, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            oWB.Close();
            oXL.Quit();

            if (mainbrowser != null) { mainbrowser.Dispose(); }
            //if (popwindow != null) { popwindow.Dispose(); }

        }

        public void StateSelect()
        {

            waitForElement("//a[@id='aet-change-loc']");
            mainbrowser.FindElement(By.XPath("//a[@id='aet-change-loc']")).Click();

            waitForElement("//input[@id='zip1']");
            mainbrowser.FindElement(By.XPath("//input[@id='zip1']")).Clear();
            System.Threading.Thread.Sleep(500); // 0.5 seconds
            mainbrowser.FindElement(By.XPath("//input[@id='zip1']")).SendKeys("Vermont");
            System.Threading.Thread.Sleep(1000); // 0.5 seconds
            mainbrowser.FindElement(By.XPath("//input[@id='zip1']")).SendKeys(Keys.Tab);
            mainbrowser.FindElement(By.XPath("//button[contains(text(),'Update my view')]")).Click();

            waitForElement("//a[contains(text(),'Dental Care')]");
            mainbrowser.FindElement(By.XPath("//a[contains(text(),'Dental Care')]")).Click();

            waitForElement("//a[contains(text(),'All Dental Professionals')]");
            mainbrowser.FindElement(By.XPath("//a[contains(text(),'All Dental Professionals')]")).Click();
            waitForLoad();

            // change location

            foreach (State USState in ListOfStates)
            //foreach (string state in stateList)
            {
                if (USState.Complete) { continue; } // if the state is done, skip this iteration

                string state = USState.Name;
                waitForElement("//a[@id='aet-change-loc']");
                mainbrowser.FindElement(By.XPath("//a[@id='aet-change-loc']")).Click();

                waitForElement("//input[@id='zip1']");
                mainbrowser.FindElement(By.XPath("//input[@id='zip1']")).Clear();
                System.Threading.Thread.Sleep(500); // 0.5 seconds
                mainbrowser.FindElement(By.XPath("//input[@id='zip1']")).SendKeys(state);
                System.Threading.Thread.Sleep(1000); // 0.5 seconds
                mainbrowser.FindElement(By.XPath("//input[@id='zip1']")).SendKeys(Keys.Tab);
                mainbrowser.FindElement(By.XPath("//button[contains(text(),'Update my view')]")).Click();
                waitForLoad();

                // Loop through Alphabet

                string AlphaRowXPath = "//div[@class='gridAlphaSort clearFix row mobilehide']/div[@class='alphasort_block']";
                StringCollection alphas2parse = new StringCollection();

                // store letters in collection since leaving this page will lose the elements and place in the list
                foreach (IWebElement alpha in mainbrowser.FindElements(By.XPath(AlphaRowXPath)))
                {
                    alphas2parse.Add(alpha.Text);
                }

                string LetterParseList = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"; // list of acceptable alphas to parse
                string NextLetter2Parse = USState.NextAlpha; // where do we start?

                // remove letters already parsed so it will skip them just like those that don't exist "*"
                if (NextLetter2Parse != "") {
                    while (LetterParseList.Substring(0,1) != NextLetter2Parse) 
                    { LetterParseList = LetterParseList.Substring(1, LetterParseList.Length - 1); }
                }

                // track position of letter in collection
                int alphalocation = 0;
                foreach (string alpha in alphas2parse)
                {                   
                    if (LetterParseList.Contains(alpha)) // skip if not a letter
                    {
                        // click letter
                        mainbrowser.FindElements(By.XPath(AlphaRowXPath))[alphalocation].Click();  // ref full path of element
                        waitForLoad();
                        StateParse(state, alpha);

                        // click reset
                        mainbrowser.FindElements(By.XPath("//button[@name='reset']"))[0].Click();
                        waitForLoad();
                    }
                    alphalocation++;
                }
            }

        }

        public void StateParse(string state, string alpha)
        {

            // Check for Next Button
            int currentPage = 1;
            Boolean NextPageExists = true;
            // Loop through additional pages if necessary

            while (NextPageExists)
            {
                // Get Jobs
                waitForLoad();
                Console.WriteLine(string.Format("--------------------- Parsing {0} - Alpha {1}-{2} ---------------------", state, alpha, currentPage.ToString()));

                // parse page
                string JobRowXPath = "//div[@class='col-xs-12 pad0 dataGridRow customPurpleRecord']";
                waitForElement(JobRowXPath);
                foreach (IWebElement jobrow in mainbrowser.FindElements(By.XPath(JobRowXPath)))
                {
                    // Pull job row for data grab
                    Dentist drec = new Dentist();

                    string CityStZip = "";
                    // Rip Data
                    try { 
                        drec.Name = jobrow.FindElement(By.ClassName("providerNameDetails")).Text.Replace("\r", "").Split('\n')[1];
                        drec.FName = drec.Name.Split(',')[1];
                        drec.LName = drec.Name.Split(',')[0];                    
                    } catch { }

                    int CSZline = 0;
                    string[] RecDetails = { };
                    RecDetails = jobrow.FindElement(By.ClassName("customDisplayForAddress")).Text.Replace("\r", "").Split('\n');

                    // find city,st zip line
                    string pattern = @",\s(AL|AK|AS|AZ|AR|CA|CO|CT|DE|DC|FM|FL|GA|GU|HI|ID|IL|IN|IA|KS|KY|LA|ME|MH|MD|MA|MI|MN|MS|MO|MT|NE|NV|NH|NJ|NM|NY|NC|ND|MP|OH|OK|OR|PW|PA|PR|RI|SC|SD|TN|TX|UT|VT|VI|VA|WA|WV|WI|WY)\s\d+";
                    foreach (string line in RecDetails)
                    {
                        if (Regex.IsMatch(line, pattern)) { break; } else { CSZline += 1; }
                    }

                    try {
                        if (CSZline > 3) {
                            drec.Address = RecDetails[CSZline-2];
                            drec.Address2 = RecDetails[CSZline-1];
                        }
                        else {
                            drec.Address = RecDetails[CSZline - 1];
                        }

                        CityStZip = RecDetails[CSZline];                     
                    } catch { }
                    
                    try { drec.City = CityStZip.Split(',')[0]; } catch { }
                    try { drec.St = CityStZip.Split(',')[1].Trim().Split(' ')[0]; } catch { }
                    try { drec.Zip = CityStZip.Split(',')[1].Trim().Split(' ')[1]; } catch { }


                    try {
                        int phoneline = 0;
                        foreach (string recline in jobrow.Text.Replace("\r", "").Split('\n'))
                        {
                            if (recline.ToUpper().Contains("PHONE NUMBER")) { phoneline++;  break; }
                            phoneline++;
                        }
                        drec.Phone = jobrow.Text.Replace("\r", "").Split('\n')[phoneline];
                    }
                    catch { }


                    try {
                        foreach (string recline in jobrow.Text.Replace("\r", "").Split('\n'))
                        {
                            if (recline.Contains("Specialties")) { drec.Specialty = recline.Split(':')[1].Trim(); }
                        }
                        //drec.Specialty = jobrow.FindElement(By.XPath("//span[contains(text(),'Specialties:')]/following-sibling::span")).Text;                     
                    } catch { }

                    // Place in Excel Row
                    //oXL.Cells[excelRow, 1] = drec.CitySection;
                    oXL.Cells[excelRow, 2] = drec.Name;
                    oXL.Cells[excelRow, 3] = drec.FName;
                    oXL.Cells[excelRow, 4] = drec.LName;
                    oXL.Cells[excelRow, 5] = drec.Address;
                    oXL.Cells[excelRow, 6] = drec.Address2;
                    oXL.Cells[excelRow, 7] = drec.City;
                    oXL.Cells[excelRow, 8] = drec.St;
                    oXL.Cells[excelRow, 9] = drec.Zip;
                    oXL.Cells[excelRow, 10] = drec.Phone;
                    //oXL.Cells[excelRow, 11] = drec.Fax;
                    oXL.Cells[excelRow, 12] = drec.Specialty;
                    //oXL.Cells[excelRow, 13] = drec.URL;

                    excelRow += 1;
                }

                try
                {
                    // Check for Next Button
                    NextPageExists = mainbrowser.FindElements(By.XPath("//a[@id='pagNext']")).Count > 0;

                    // Click Next page button
                    if (NextPageExists)
                    {
                        mainbrowser.FindElement(By.XPath("//a[@class='next-link']")).Click();
                    }

                    currentPage++;
                    if (currentPage%10==0) {
                        System.Threading.Thread.Sleep(180*1000); // 120 seconds
                    }
                                                                      //URL = mainbrowser.Url;

                    }
                catch { }


            }

        }

        public void waitForLoad()
        {

            int currentWait = 0;
            int waitIncrement = 1000;
            int maxWait = 60 * 1000;
            try
            {
                while (mainbrowser.FindElements(By.XPath("//div[@id='newNavSpinner2']")).Count > 0 && (currentWait <= maxWait))
                {
                    System.Threading.Thread.Sleep(waitIncrement); // 0.5 seconds
                    currentWait += waitIncrement;
                    Console.WriteLine("...Loading Spinner Wait Time: " + currentWait.ToString());
                }

            }
            catch (Exception ex) { }

        }

        public void waitForElement(string Xpath)
        {

            int currentWait = 0;
            int waitIncrement = 1000;
            int maxWait = 60 * 1000;
            try
            {
                while (mainbrowser.FindElements(By.XPath(Xpath)).Count < 1 && (currentWait <= maxWait))
                {
                    System.Threading.Thread.Sleep(waitIncrement); // 0.5 seconds
                    currentWait += waitIncrement;
                    Console.WriteLine("...Element Wait Time: " + currentWait.ToString());
                }

            }
            catch (Exception ex) { }

        }

    }
}
