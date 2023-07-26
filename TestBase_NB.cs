using NUnit.Framework;

using OpenQA.Selenium;

using OpenQA.Selenium.Chrome;

using System.Threading;

using OpenQA.Selenium.Support.UI;

using System.IO;

using System;

using OpenQA.Selenium.Firefox;
using System.Data.OleDb;
using System.Data;
using System.Collections.Generic;

namespace ILR_TestSuite

{

    [TestFixture]

    public class TestBase_NB

    {

        private ChromeOptions _chromeOptions;

        public IWebDriver _driver;

        private string _userName;

        private string _password;

        private string _pin;

        public string _screenShotFolder;

        public string _test_data_connString;
        public string _test_results_connString;
        private string _gcode = "G992127";


        [SetUp]

        public void StartBrowser()
        {


            _chromeOptions = new ChromeOptions();
            _chromeOptions.AddArguments("--incognito");
            _chromeOptions.AddArguments("--ignore-certificate-errors");
            _driver = new ChromeDriver("C:/Code/bin", _chromeOptions);

            _test_data_connString = "Provider= Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:/Users/G992127/Documents/GitHub/ILR_TestSuite/ILR_TestSuite/New Business/SalesAppBase/TestData.xlsx" + ";Extended Properties='Excel 8.0;HDR=Yes'";
            _test_results_connString = "Provider= Microsoft.ACE.OLEDB.12.0;" + "Data Source=C/Users/G992127/Documents/GitHub/ILR_TestSuite/ILR_TestSuite/New Business/SalesAppBase/TestResults.xlsx" + ";Extended Properties='Excel 8.0;HDR=Yes'";
            _screenShotFolder = $@"C:\Users\{_gcode}\Documents\GitHub\ILR_TestSuite\ILR_TestSuite\New Business\​{ScreenShotDailyFolderName()}​\";


            new DirectoryInfo(_screenShotFolder).Create();


        }

        private static string ScreenShotDailyFolderName()

        {

            return DateTime.Now.ToString("yyyyMMdd").Replace("AM", string.Empty).Replace("PM", string.Empty);

        }

        public void TakeScreenshot(IWebDriver driver, string filePath, string fileName)

        {

            if (!Directory.Exists(filePath))

                new DirectoryInfo(filePath).Create();


            ITakesScreenshot ssdriver = driver as ITakesScreenshot;

            Screenshot screenshot = ssdriver.GetScreenshot();

            fileName = $"{fileName}{ScreenShotTime()}.png";

            screenshot.SaveAsFile($"{filePath}{fileName}", ScreenshotImageFormat.Png);

            // byte[] byteArray = ((ITakesScreenshot)driver).GetScreenshot().AsByteArray;

            //Bitmap ss  = new Bitmap(new System.IO.MemoryStream(byteArray));

            // screenshot.SaveAsFile(String.Format(fileName, ImageFormat.Png));

        }

        private static string ScreenShotTime()

        {

            return DateTime.Now.TimeOfDay.ToString().Replace(":", "_").Replace(".", "_");

        }



        public IWebDriver SiteConnection()

        {

            _driver.Url = "https://uat-fe.safricansalesapp.net/advisor/dashboard/";

            try
            {

               

                _userName = "G992092";//TODO add your user name and password

                _password = "SCHggy024588";

                _pin = "880808";

                _driver.Manage().Window.Maximize();

                System.Threading.Thread.Sleep(2000);

                _driver.Manage().Window.Maximize();

                System.Threading.Thread.Sleep(2000);

                _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/div/div[2]/button")).Click();
                System.Threading.Thread.Sleep(4000);

                _driver.SwitchTo().Frame("form-frame");
                System.Threading.Thread.Sleep(2000);
                IWebElement loginTextBox = _driver.FindElement(By.Name("username"));
               
               
                System.Threading.Thread.Sleep(3000);
                IWebElement passwordTextBox = _driver.FindElement(By.Name("password"));
                System.Threading.Thread.Sleep(3000);
                IWebElement loginBtn = _driver.FindElement(By.Name("action"));
                System.Threading.Thread.Sleep(3000);
                loginTextBox.SendKeys(_userName);
                System.Threading.Thread.Sleep(3000);
                passwordTextBox.SendKeys(_password);
                System.Threading.Thread.Sleep(3000);
                loginBtn.Click();
                System.Threading.Thread.Sleep(3000);
               //_driver.SwitchTo().DefaultContent();

                Delay(10);


                //create pin
                var pin1 = _driver.FindElement(By.Name("pin"));
                pin1.SendKeys(_pin);
                System.Threading.Thread.Sleep(4000);
                _driver.FindElement(By.Name("pinConfirm")).SendKeys(_pin);
                System.Threading.Thread.Sleep(4000);
                IWebElement create = _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div/section/form/button"));
                create.Click();
                Delay(2);
                

                return _driver;
            }
            catch (Exception ex) {
                 throw ex;
            }

           
           
        }
        public string GetPolicyNoFromExcell(string sheet, string function)
        {


            string conRef = "";
            using (OleDbConnection conn = new OleDbConnection(_test_data_connString))
            {
                try
                {

                    // Open connection
                    conn.Open();
                    string cmdQuery = "SELECT * FROM [" + sheet + "$]";

                    OleDbCommand cmd = new OleDbCommand(cmdQuery, conn);

                    // Create new OleDbDataAdapter
                    OleDbDataAdapter oleda = new OleDbDataAdapter();

                    oleda.SelectCommand = cmd;

                    // Create a DataSet which will hold the data extracted from the worksheet.
                    DataSet ds = new DataSet();

                    // Fill the DataSet from the data extracted from the worksheet.
                    oleda.Fill(ds, "Policies");



                    foreach (var row in ds.Tables[0].DefaultView)
                    {

                        var func = ((System.Data.DataRowView)row).Row.ItemArray[1].ToString();

                        if (func == function)
                        {
                            conRef = ((System.Data.DataRowView)row).Row.ItemArray[0].ToString();
                        }
                    }




                }
                catch (Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    conn.Close();


                    conn.Dispose();
                }
            }

            return conRef;

        }
        public Decimal getPremuimFromRateTable(string age, string rolePlayer, string sumAsured, string product)
        {
            var premium = "";
            var connectionStr = "Provider= Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:/Users/G992107/Documents/GitHub/ILR_TestSuite/ILR_TestSuite/MIP UAT Test Scenarios/TestData.xlsx" + ";Extended Properties='Excel 8.0;HDR=Yes'";

            using (OleDbConnection conn = new OleDbConnection(connectionStr))
            {
                try
                {

                    var cover = rolePlayer + "_" + sumAsured;
                    // Open connection
                    conn.Open();
                    string cmdQuery = $"SELECT Term, {cover} FROM [{product}$]";

                    OleDbCommand cmd = new OleDbCommand(cmdQuery, conn);

                    // Create new OleDbDataAdapter
                    OleDbDataAdapter oleda = new OleDbDataAdapter();

                    oleda.SelectCommand = cmd;

                    // Create a DataSet which will hold the data extracted from the worksheet.
                    DataSet ds = new DataSet();

                    // Fill the DataSet from the data extracted from the worksheet.
                    oleda.Fill(ds, "Policies");



                    foreach (var row in ds.Tables[0].DefaultView)
                    {

                        var excellAge = ((System.Data.DataRowView)row).Row.ItemArray[0].ToString();
                        if (excellAge == age)
                        {
                            premium = ((System.Data.DataRowView)row).Row.ItemArray[1].ToString();
                        }


                    }

                }
                catch (Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                }


            }
            return Convert.ToDecimal(premium);




        }
        public void writeResultsToExcell(Tuple<string,string> results, string sheet, string scenario)
        {
            using (OleDbConnection conn = new OleDbConnection(_test_data_connString))
            {
                try
                {

                    // Open connection
                    conn.Open();
                    OleDbCommand cmd = conn.CreateCommand();

                    var testDate = DateTime.Now.ToString();

                    //Test_Date
                  
                    cmd.CommandText = $"UPDATE [{sheet}$] SET Test_Results  = '{results.Item1}' WHERE Scenario_ID = '{scenario}';";
                    cmd.ExecuteNonQuery();
                    cmd.CommandText = $"UPDATE [{sheet}$] SET Comment  = '{results.Item2}' WHERE Scenario_ID = '{scenario}';";
                    cmd.ExecuteNonQuery();
                    cmd.CommandText = $"UPDATE [{sheet}$] SET Test_Date = '{testDate}' WHERE Scenario_ID = '{scenario}';";
                    cmd.ExecuteNonQuery();

                }
                catch (Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
        }
        public Dictionary<string, List<Dictionary<string, string>>> getRolePlayers(string id)
        {
 
           Dictionary < string, List<Dictionary<string,string>>> roleplayersData = new Dictionary<string, List<Dictionary<string,string>>>();
            var sheets = new List<string>();
            sheets.Add("PolicyHolder_Details");
            sheets.Add("Extended");
            sheets.Add("Parents");
            sheets.Add("Children");
            sheets.Add("spouse");
            sheets.Add("PolicyPayer");
            sheets.Add("Beneficiaries");

            foreach (var sheet in sheets)
            {
                List<Dictionary<string, string>> roleplyers = new List<Dictionary<string, string>>();
                List<string> keys = new List<string>();
                //Loop through every row of that sheet and get data of all the players with the scenario id given
                _test_data_connString = "Provider= Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:/Users/G992107/Documents/GitHub/ILR_TestSuite/ILR_TestSuite/New Business/SalesAppBase/TestData.xlsx" + ";Extended Properties='Excel 8.0;HDR=NO'";
                using (OleDbConnection conn = new OleDbConnection(_test_data_connString))
                {
                    try
                    {
                        // Open connection
                        conn.Open();
                        string cmdQuery = "SELECT * FROM [" + sheet + "$]";

                        OleDbCommand cmd = new OleDbCommand(cmdQuery, conn);

                        // Create new OleDbDataAdapter
                        OleDbDataAdapter oleda = new OleDbDataAdapter();

                        oleda.SelectCommand = cmd;

                        // Create a DataSet which will hold the data extracted from the worksheet.
                        DataSet ds = new DataSet();

                        // Fill the DataSet from the data extracted from the worksheet.
                        oleda.Fill(ds, "PolicyHolderData");

                        var count = 0;
                        
                        int noColumn = ds.Tables[0].Columns.Count;
                        //Loop throgh every row
                        foreach (var row in ds.Tables[0].DefaultView)
                        {
                            if (count == 0)
                            {
                                for (int i = 1; i < noColumn; i++)
                                {

                                    keys.Add(((System.Data.DataRowView)row).Row.ItemArray[i].ToString());

                                }

                            }
                            else
                            {
                                var scenarioID = ((System.Data.DataRowView)row).Row.ItemArray[0].ToString();


                                if (scenarioID == id)
                                {

                                    var column = 1;
                                    Dictionary<string, string> rowData = new Dictionary<string, string>();

                                    foreach (string key in keys)
                                   {
                                        rowData.Add(key, ((System.Data.DataRowView)row).Row.ItemArray[column].ToString());
                                       column++;
                                   }
                                    roleplyers.Add(rowData);

                                }
                            }
                            count++;

                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                    finally
                    {
                        conn.Close();


                        conn.Dispose();
                    }
                }
                roleplayersData.Add(sheet, roleplyers);
            }

            return roleplayersData;

        }
        public Dictionary<string, string> getPolicyHolderDetails(string scenario_id)
        {
            _test_data_connString = "Provider= Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:/Users/G992107/Documents/GitHub/ILR_TestSuite/ILR_TestSuite/New Business/SalesAppBase/TestData.xlsx" + ";Extended Properties='Excel 8.0;HDR=NO'";
            //Variables to store policy holder data
            Dictionary<string, string> policyHolderData = new Dictionary<string, string>();
       
        
            //Sheets in the test data file that we want to access to extract Policy holder data
            var sheets = new List<string>();
            sheets.Add("PolicyHolder_Details");
            sheets.Add("Affordability_Check");
            sheets.Add("BankDetails");
            sheets.Add("PhysicalAddress");
            //Extracting data from every sheet 
            
            foreach (var sheet in sheets)
            {
               
                using (OleDbConnection conn = new OleDbConnection(_test_data_connString))
                {
                    try
                    {
                        // Open connection
                        conn.Open();
                        string cmdQuery = "SELECT * FROM [" + sheet + "$]";

                        OleDbCommand cmd = new OleDbCommand(cmdQuery, conn);

                        // Create new OleDbDataAdapter
                        OleDbDataAdapter oleda = new OleDbDataAdapter();

                        oleda.SelectCommand = cmd;

                        // Create a DataSet which will hold the data extracted from the worksheet.
                        DataSet ds = new DataSet();

                        // Fill the DataSet from the data extracted from the worksheet.
                        oleda.Fill(ds, "PolicyHolderData");

                        var count = 0;
                        List<string> keys = new List<string>();
                        int noColumn = ds.Tables[0].Columns.Count;
                        foreach (var row in ds.Tables[0].DefaultView)
                        {
                           
                            if (count == 0)
                            {
                                for (int i = 1; i < noColumn ; i++)
                                {
                                    
                                        keys.Add(((System.Data.DataRowView)row).Row.ItemArray[i].ToString());
                              
                                }

                            }
                            else
                            {
                                var scenarioID = ((System.Data.DataRowView)row).Row.ItemArray[0].ToString();


                                if (scenarioID == scenario_id)
                                {
                                   
                                    var column = 1;


                                    foreach (string key in keys)
                                    {
                                        policyHolderData.Add(key, ((System.Data.DataRowView)row).Row.ItemArray[column].ToString());
                                        column++;
                                    }

                                }
                            }
                            count++;
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                    finally
                    {
                        conn.Close();


                        conn.Dispose();
                    }
                }

            }

            return policyHolderData;
        }

        public void DisconnectBrowser()

        {

            if (_driver != null)

                _driver.Quit();

        }



        public string GetSeleniumFormatTag(string inputControlName)

        {

            var result = $"//*[@id=\"{inputControlName}\"]";

            return result;

        }



        public void Delay(int delaySeconds)

        {

            Thread.Sleep(delaySeconds * 1000);

        }



    }

}