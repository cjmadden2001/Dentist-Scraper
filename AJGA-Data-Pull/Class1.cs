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


namespace Dentist_Scraper
{

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

            var outputFile = new FileInfo(Path.Combine(docPath, "Dentists.xlsx"));
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
            StateSelect("http://www.nationaldirectoryofdentists.com/dentists/");
            for (int i = 0; i < cities.Count; i++)
            {
                CityParse(cities[i]);
            }

            //CityParse("http://www.nationaldirectoryofdentists.com/dentists/AL/Decatur");
            //CityParse("http://www.nationaldirectoryofdentists.com/dentists/AL/Birmingham");

            oWB.SaveAs(outputFile, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            oWB.Close();
            oXL.Quit();

            if (mainbrowser != null) { mainbrowser.Dispose(); }
            if (popwindow != null) { popwindow.Dispose(); }

        }
        
        public void StateSelect(string mainURL)
        {
            mainbrowser.Navigate().GoToUrl(mainURL);

            // parse page
            string stateXPath = "//div[@id='content']/div/div/div/p";
            foreach (IWebElement state in mainbrowser.FindElements(By.XPath(stateXPath)))
            {
                string stateList = "Tennessee,Texas,Utah,Vermont,Virginia,Washington,West Virginia,Wisconsin,Wyoming";
                // Indiana,Iowa,Kansas,Kentucky,Louisiana,Maine,Maryland,Massachusetts,Michigan,Minnesota,Mississippi,Missouri,Montana,Nebraska,Nevada,New Hampshire, New Jersey,New Mexico, New York,North Carolina, North Dakota,Ohio,Oklahoma,Oregon,Pennsylvania,Rhode Island, South Carolina,South Dakota,
                //"Alabama,Alaska,Arizona,Arkansas,California,Colorado,Connecticut,Delaware,District Of Columbia,Florida,Georgia,Hawaii,Idaho,Illinois,";
                if (stateList.Contains(state.Text))
                {
                    string stateURL = state.FindElement(By.TagName("a")).GetAttribute("href");
                    Console.WriteLine(string.Format("--------------------- Pulling {0} Cities ---------------------", state.FindElement(By.TagName("a")).Text));
                    popwindow.Navigate().GoToUrl(stateURL);
                    CityList(stateURL);
                }
            }
            System.Threading.Thread.Sleep(1 * 1000); // wait one second after each state
        }

        public void CityList(string stateURL)
        {
            // parse page
            string cityXPath = "//div[@id='content']/div/div/div/p";
            foreach (IWebElement city in popwindow.FindElements(By.XPath(cityXPath)))
            {
                string cityURL = null;
                    try { 
                        cityURL = city.FindElement(By.TagName("a")).GetAttribute("href");
                    }
                    catch {
                        Console.WriteLine(string.Format("-**** ERROR ON CITY:  {0}", cityURL));
                        //cityURL = city.FindElement(By.TagName("a")).GetAttribute("href");
                }
                cities.Add(cityURL);
            }
        }
        
        public void CityParse(string CityURL)
        {
            mainbrowser.Navigate().GoToUrl(CityURL);

            // Check for Next Button
            int currentPage = 1;
            Boolean NextPageExists = true;
            // Loop through additional pages if necessary

            while (NextPageExists)
            {
                // Get Jobs
                Console.WriteLine(string.Format("--------------------- Parsing {0} - {1}", CityURL, DateTime.Now.ToString("G")));

                // parse page
                string JobRowXPath = "//ul[@class='results']/li";

                foreach (IWebElement jobrow in mainbrowser.FindElements(By.XPath(JobRowXPath)))
                {
                    // Pull job row for data grab
                    Dentist drec = new Dentist();

                    // Get URL for Dentist Record and pop in other window
                    string popURL = jobrow.FindElement(By.TagName("a")).GetAttribute("href");
                    popwindow.Navigate().GoToUrl(popURL);

                    string[] RecDetails = { };
                    string CityStZip = ""; string AddressLine = "";
                    // Rip Data
                    try { drec.CitySection = CityURL; } catch { }
                    try { drec.Name = popwindow.FindElement(By.XPath("//div[@class='prodetailsname']")).Text; } catch { }
                    try { drec.FName = drec.Name.Split(' ')[0]; } catch { }
                    try {
                        int pieceNum = 1;

                        string[] NamePieces = drec.Name.Substring(drec.FName.Length + 1).Split(' ');
                        string lastPiece = NamePieces[NamePieces.Count() - 1];
                        if (lastPiece == "DMD" || lastPiece == "DDS") {
                            NamePieces = NamePieces.Where(w => w != lastPiece).ToArray();
                        }
                        int totalPieces = NamePieces.Count();

                        foreach (string piece in NamePieces)
                        {
                            Boolean IsMiddle = (pieceNum < totalPieces);
                            
                            if (IsMiddle && piece.Length > 1 && piece.Length<3)
                            {
                                drec.LName += piece;
                            }

                            if (!IsMiddle)
                            {
                                drec.LName += piece;
                            }

                            if (IsMiddle && (piece.Length == 1 || piece.Length > 2)) {
                                drec.FName += " " + piece;
                            }
                            pieceNum += 1;
                        }
                    
                    } catch { }

                    int CSZline = 0;
                    
                    // Grab the capture location so we can select correct city, st when multiple exist in address fields
                    string captureLoc = popwindow.FindElement(By.XPath("//div[@id='resultsheader']")).Text;
                    captureLoc = captureLoc.Substring(captureLoc.IndexOf("Dentists in") + 12);

                    try
                    {
                        // create rec details array
                        foreach (IWebElement addressrow in popwindow.FindElements(By.XPath("//div[@class='address']")))
                        {
                            if (addressrow.Text.Contains(captureLoc))
                            {
                                RecDetails = addressrow.Text.Replace("\r", "").Split('\n');
                                break;
                            }
                        }
                        // check for Business Header line, rip it out
                        if (popwindow.FindElement(By.XPath("//div[@class='address']")).GetAttribute("innerHTML").Contains("<b>"))
                        { RecDetails = RecDetails.Where(w => w != RecDetails[0]).ToArray(); }

                        // find city,st zip line
                        string pattern = @",\s(AL|AK|AS|AZ|AR|CA|CO|CT|DE|DC|FM|FL|GA|GU|HI|ID|IL|IN|IA|KS|KY|LA|ME|MH|MD|MA|MI|MN|MS|MO|MT|NE|NV|NH|NJ|NM|NY|NC|ND|MP|OH|OK|OR|PW|PA|PR|RI|SC|SD|TN|TX|UT|VT|VI|VA|WA|WV|WI|WY)\s\d+";
                        foreach (string line in RecDetails)
                        { 
                            if (Regex.IsMatch(line, pattern)) { break; } else { CSZline += 1; }
                        }

                        if (RecDetails.Count() > 1) 
                            { 
                            AddressLine = RecDetails[0];                            
                            CityStZip = RecDetails[CSZline]; 
                        } 
                            else { CityStZip = RecDetails[0]; }
                    } catch { }
                    try {
                        // double address line exists
                        if (CSZline == 2)
                        {
                            drec.Address = RecDetails[0];
                            drec.Address2 = RecDetails[1];
                        }
                        else
                        {
                            // single address line exists
                            if (CSZline == 1) { 
                                // # 123 at the end
                                if (AddressLine.Contains(" #"))
                                {
                                    drec.Address = AddressLine.Substring(0, AddressLine.IndexOf(" #"));
                                    drec.Address2 = AddressLine.Substring(AddressLine.IndexOf(" #") + 1);
                                }
                                // Ste 123 at the end
                                else if (AddressLine.ToUpper().Contains(" STE"))
                                {
                                    drec.Address = AddressLine.Substring(0, AddressLine.ToUpper().IndexOf(" STE"));
                                    drec.Address2 = AddressLine.Substring(AddressLine.ToUpper().IndexOf(" STE") + 1);
                                }
                                // Suite 123 at the end
                                else if (AddressLine.ToUpper().Contains(" SUITE"))
                                {
                                    drec.Address = AddressLine.Substring(0, AddressLine.ToUpper().IndexOf(" SUITE"));
                                    drec.Address2 = AddressLine.Substring(AddressLine.ToUpper().IndexOf(" SUITE") + 1);
                                }
                                else if (AddressLine.ToUpper().Substring(0, 5) == "SUITE")
                                {
                                    AddressLine = AddressLine.Substring(AddressLine.IndexOf(" ") + 1);
                                    drec.Address = AddressLine.Substring(AddressLine.IndexOf(" ") + 1);
                                    drec.Address2 = "Suite " + AddressLine.Split(' ')[0];
                                }
                                else { drec.Address = AddressLine; }
                            }
                        }

                    } catch { }
                    try { drec.City = CityStZip.Split(',')[0]; } catch { }
                    try { drec.St = CityStZip.Split(',')[1].Trim().Split(' ')[0]; } catch { }
                    try { drec.Zip = CityStZip.Split(',')[1].Trim().Split(' ')[1]; } catch { }

                    for (int i = CSZline + 1; i < RecDetails.Count(); i++)
                    {
                        if (RecDetails[i].Contains("phone")) { drec.Phone = RecDetails[i].Split(':')[1].Trim(); }
                        if (RecDetails[i].Contains("fax")) { drec.Fax = RecDetails[i].Split(':')[1].Trim(); }
                    }

                    try { drec.Specialty = popwindow.FindElement(By.XPath("//div[@class='prodetailsspecialty']")).Text; } catch { }
                    try { drec.URL = popURL; } catch { }

                    // Place in Excel Row
                    oXL.Cells[excelRow, 1] = drec.CitySection;
                    oXL.Cells[excelRow, 2] = drec.Name;
                    oXL.Cells[excelRow, 3] = drec.FName;
                    oXL.Cells[excelRow, 4] = drec.LName;
                    oXL.Cells[excelRow, 5] = drec.Address;
                    oXL.Cells[excelRow, 6] = drec.Address2;
                    oXL.Cells[excelRow, 7] = drec.City;
                    oXL.Cells[excelRow, 8] = drec.St;
                    oXL.Cells[excelRow, 9] = drec.Zip;
                    oXL.Cells[excelRow, 10] = drec.Phone;
                    oXL.Cells[excelRow, 11] = drec.Fax;
                    oXL.Cells[excelRow, 12] = drec.Specialty;
                    oXL.Cells[excelRow, 13] = drec.URL;

                    excelRow += 1;
                }

                try
                {
                    // Check for Next Button
                    NextPageExists = mainbrowser.FindElements(By.XPath("//a[contains(text(),'next >')]")).Count > 0;

                    // Click Next page button
                    if (NextPageExists)
                    {
                        mainbrowser.FindElement(By.XPath("//a[contains(text(),'next >')]")).Click();
                    }

                    currentPage++;
                    CityURL = mainbrowser.Url;

                }
                catch { }


            }

        }

    }
}
