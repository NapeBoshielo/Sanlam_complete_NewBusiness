using ILR_TestSuite;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using OpenQA.Selenium.Firefox;
using System.Data.OleDb;
using System.Data;
using OpenQA.Selenium.Interactions;


namespace ILR_TestSuite.New_Business.Sales_App
{

    [TestFixture]
    public class SalesApp : TestBase_NB

    {
      
        [SetUp]
            public void startBrowser()

            {

                _driver = base.SiteConnection();
           

            }

            [Test, Order(1)]
        [Obsolete]
        public void RunTest()
            {
                Delay(15);
                          
                
            using (OleDbConnection conn = new OleDbConnection(_test_data_connString))
                {
                    try
                    {
                        var sheet = "Scenarios";
                        // Open connection
                        conn.Open();
                        string cmdQuery = "SELECT * FROM ["+ sheet + "$]";

                        OleDbCommand cmd = new OleDbCommand(cmdQuery, conn);

                        // Create new OleDbDataAdapter
                        OleDbDataAdapter oleda = new OleDbDataAdapter();

                        oleda.SelectCommand = cmd;

                        // Create a DataSet which will hold the data extracted from the worksheet.
                        DataSet ds = new DataSet();

                        // Fill the DataSet from the data extracted from the worksheet.
                        oleda.Fill(ds, "Policies");


                        //addMainLife();
                        foreach (var row in ds.Tables[0].DefaultView)
                        {
                            var Scenario_ID = ((System.Data.DataRowView)row).Row.ItemArray[0].ToString();
                            var results = PositiveTestProcess(Scenario_ID);

                            OleDbCommand updatecmd = conn.CreateCommand();

                            //Test_Date
                            var testDate = DateTime.Now.ToString();
                            updatecmd.CommandText = $"UPDATE [{sheet}$] SET Test_Results  = '{results.Item1}' WHERE Scenario_ID = '{Scenario_ID}';";
                            updatecmd.ExecuteNonQuery();
                            updatecmd.CommandText = $"UPDATE [{sheet}$] SET Comment  = '{results.Item2}' WHERE Scenario_ID = '{Scenario_ID}';";
                            updatecmd.ExecuteNonQuery();
                            updatecmd.CommandText = $"UPDATE [{sheet}$] SET Test_Date = '{testDate}' WHERE Scenario_ID = '{Scenario_ID}';";
                            updatecmd.ExecuteNonQuery();


                        Delay(4);
                        //Click on Menu
                        _driver.FindElement(By.XPath("/html/body/div[1]/div[1]/nav/button")).Click();
                        Delay(4);
                        //Click on Dashbaord
                        //*[@id="gatsby-focus-wrapper"]/div/section[1]/a[1]
                        _driver.FindElement(By.XPath("/html/body/div[1]/div[1]/div[1]/section[1]/a[1]")).Click();



                    }
                }
                    finally
                    {
                        conn.Close();
                        conn.Dispose();
                    }
                }

            }

        [Obsolete]
        public Tuple<string, string> PositiveTestProcess(string scenario_ID)
        {
            var upload_file = "C:/Users/G992107/Documents/GitHub/ILR_TestSuite/ILR_TestSuite/New Business/upload/download.jpg";
            Delay(10);
            string results = "", comment = "";

            //get policy holder data
            var policyHolderData = getPolicyHolderDetails(scenario_ID);
            _driver.SwitchTo().ActiveElement();
            _driver.FindElement(By.XPath("//*[@id='___gatsby']"));
            Delay(20);
            IWebElement new_client = _driver.FindElement(By.ClassName("new-app-button"));
            new_client.Click();
            //  Actions action = new Actions(_driver);
            // action.MoveToElement(new_client).Perform()
            Delay(2);
            IWebElement town = _driver.FindElement(By.Name("town"));
            town.SendKeys(policyHolderData["Town"]);
            Delay(1);
            town.SendKeys(Keys.ArrowDown);
            Delay(1);
            town.SendKeys(Keys.Enter);
            Delay(4);
            IWebElement worksite = _driver.FindElement(By.Name("worksite"));
            worksite.SendKeys(policyHolderData["Worksite"]);
            Delay(1);
            worksite.SendKeys(Keys.ArrowDown);
            Delay(1);
            worksite.SendKeys(Keys.Enter);
            Delay(2);
            IWebElement employer = _driver.FindElement(By.Name("employer-name"));
            employer.SendKeys(policyHolderData["Employment"]);
            Delay(1);
            employer.SendKeys(Keys.ArrowDown);
            Delay(1);
            employer.SendKeys(Keys.Enter);
            Delay(2);
            IWebElement yes = _driver.FindElement(By.XPath("/html/body/reach-portal/div/div/div/div[4]/div/div[2]/div/label[1]"));
            Delay(1);
            yes.Click();
            Delay(2);
            IWebElement cont = _driver.FindElement(By.XPath(" /html/body/reach-portal/div/div/div/div[5]/button[2]"));
            cont.Click();
            Delay(4);
            IWebElement agree = _driver.FindElement(By.XPath(" /html/body/reach-portal/div/div/div/div[2]/button"));
            agree.Click();
            //Personal Details
            Delay(2);
            //firstname
            _driver.FindElement(By.XPath("//*[@id='/name']")).SendKeys(policyHolderData["First_name"]);
            Delay(2);
            //maiden name
            _driver.FindElement(By.XPath("//*[@id='/maiden-surname']")).SendKeys(policyHolderData["Maiden_Surname"]);
            Delay(2);
            //Id 
            Delay(3);
            _driver.FindElement(By.XPath("//*[@id='/id-number']")).SendKeys(policyHolderData["ID_number"]);
            Delay(2);
            _driver.FindElement(By.XPath("//*[@id='/surname']")).SendKeys(policyHolderData["Surname"]);
            Delay(2);
            //Select ethicity
            IWebElement select = _driver.FindElement(By.XPath(" //*[@id='/ethnicity']"));
            SelectElement oselect = new SelectElement(select);
            oselect.SelectByValue(policyHolderData["Ethnicity"]);
            Delay(2);
            //Select Maratiel
            IWebElement selectstatus = _driver.FindElement(By.XPath("//*[@id='/marital-status']"));
            SelectElement cselect = new SelectElement(selectstatus);
            cselect.SelectByValue(policyHolderData["Marital_Status"]);
            Delay(2);
            //Enter contact number
            _driver.FindElement(By.XPath("//*[@id='/contact-number']")).SendKeys(policyHolderData["CellPhone_number"]);
            Delay(2);
            //Enter email
            _driver.FindElement(By.XPath("//*[@id='/email']")).SendKeys(policyHolderData["Email"]);
            Delay(2);
            //Enter gross monthly
            _driver.FindElement(By.XPath("//*[@id='/gross-monthly-income']")).SendKeys(policyHolderData["Gross"]);
            Delay(2);
            //Select employent type
            if (policyHolderData["Permanent"] == "Yes")
            {
                _driver.FindElement(By.XPath("/html/body/div[1]/div[1]/article/section/form/div/div[16]/div/label[1]")).Click();
            }
            else
            {
                _driver.FindElement(By.XPath("/html/body/div[1]/div[1]/article/section/form/div/div[16]/div/label[2]")).Click();
            }
            Delay(2);
            //Salary frequency
            switch (policyHolderData["Salary_frequency"])
            {
                case "Weekly":
                    _driver.FindElement(By.XPath("/html/body/div[1]/div[1]/article/section/form/div/div[17]/div/label[1]")).Click();
                    break;
                case "Monthly":
                    _driver.FindElement(By.XPath("/html/body/div[1]/div[1]/article/section/form/div/div[17]/div/label[2]")).Click();
                    break;
                case "Other":
                    _driver.FindElement(By.XPath("/html/body/div[1]/div[1]/article/section/form/div/div[17]/div/label[3]")).Click();
                    break;
                default:
                    _driver.FindElement(By.XPath("/html/body/div[1]/div[1]/article/section/form/div/div[17]/div/label[2]")).Click();
                    break;
            }


            //click next 
            Delay(2);
            try
            {
                _driver.FindElement(By.XPath("/html/body/div[1]/div[1]/div[2]/div[1]/a")).Click();

            }
            catch
            {
                //Age validarion 
                _driver.FindElement(By.XPath("/html/body/div[1]/div[1]/div[2]/div[1]/button")).Click();

                String ValidationMsg = _driver.FindElement(By.XPath("/html/body/div[1]/div[1]/article/section/form/div/div[8]/span/span")).Text;
                comment = "Main Life Assured ";
                comment += ValidationMsg;
                if (ValidationMsg.Contains("Must be at least 18 years old.") || ValidationMsg.Contains("Must not be older than 74 years of age."))
                {
                    results = "Failed";
                    TakeScreenshot(_driver, $@"{_screenShotFolder}\Validations\", "MainLife_Age");

                    return Tuple.Create(results, comment);

                }


            }

            //occupation
            Delay(3);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/form/section/div/div[1]/label")).Click();
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[3]/div[1]/a[2]")).Click();
            ///dependants
            Delay(4);
            for (int i = 1; i < 5; i++)
            {
                _driver.FindElement(By.XPath($"//*[@id='gatsby-focus-wrapper']/article/section/form/div[1]/section[{i.ToString()}]/label")).Click();
            }
            //click next
            Delay(3);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div[1]/a[2]")).Click();
            Delay(3);
            for (int i = 1; i < 5; i++)
            {
                _driver.FindElement(By.XPath($"//*[@id='gatsby-focus-wrapper']/article/section/form/div[1]/section[{i.ToString()}]/label/section/div[1]")).Click();
            }
            //click next
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div[1]/a[2]")).Click();
            //sclick on non applicable 
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/div[1]/div/button")).Click();
            //Net Salary After Deductions
            Delay(1);
            _driver.FindElement(By.Name("/total-salary-after-deductions")).SendKeys(policyHolderData["Net_Salary"]);
            //Additional income
            Delay(1);
            _driver.FindElement(By.Name("/additional-income")).SendKeys(policyHolderData["Additional_Income"]);
            //Existing Financial Cover
            Delay(1);
            _driver.FindElement(By.Name("/existing-financial-cover")).SendKeys(policyHolderData["Existing_FinancialCover"]);
            //School Fees
            Delay(1);
            _driver.FindElement(By.Name("/school-fees")).SendKeys(policyHolderData["School_Fees"]);
            //Food
            Delay(1);
            _driver.FindElement(By.Name("/food")).SendKeys(policyHolderData["Food"]);
            //Retail accounts
            Delay(1);
            _driver.FindElement(By.Name("/retail-accounts")).SendKeys(policyHolderData["Retail_accounts"]);
            //Cellphone
            Delay(1);
            _driver.FindElement(By.Name("/cellphone")).SendKeys(policyHolderData["Cellphone"]);
            //Debt
            Delay(1);
            _driver.FindElement(By.Name("/debt")).SendKeys(policyHolderData["Debt"]);
            // Mortgage / Rent
            Delay(1);
            _driver.FindElement(By.Name("/mortage-rent")).SendKeys(policyHolderData["Mortgage_Rent"]);
            //Transport
            Delay(1);
            _driver.FindElement(By.Name("/transport")).SendKeys(policyHolderData["Transport"]);
            //Entertainment / Other
            Delay(1);
            _driver.FindElement(By.Name("/entertainment-other")).SendKeys(policyHolderData["Entertainment_Other"]);
            //click next
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div[1]/a[2]")).Click();
            //click tickbox for agreement 
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section/div[2]/input[1]")).Click();
            //click next
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div[1]/a[2]")).Click();


            var policyplayers = getRolePlayers(scenario_ID);
            List<string> keys = new List<string>();
            keys.Add("PolicyHolder_Details");
            keys.Add("spouse");
            keys.Add("Children");
            keys.Add("Parents");
            keys.Add("Extended");

            var beneficiaries = policyplayers["Beneficiaries"];

            //click tickbox product
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/section/div[1]/div[3]/button/span")).Click();

            //click tickbox product
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/div/div[2]/label/div")).Click();


            //click next
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div[1]/a")).Click();
            //click on No
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section[1]/div/div[2]/div/div/label[2]")).Click();

            //click on 5%
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section[2]/div/div[2]/div/div/label[1]")).Click();
            //Add Provided LAs
            var lifeAsuredCounter = 0;
            var label = 1;
            var section = 3;
            IWebElement DOB;
            String date_of_birth = "", frontEndPrem = "", frontEndMin = "", frontEndMax = "";
            Tuple<string, string> validation;
           

            foreach (var key in keys)
            {
                
                foreach (var item in policyplayers[key])
                {

                    if (item.Count > 0)
                    {
                        if (key == "PolicyHolder_Details")
                        {
                            if (item["Covered"] == "Yes")
                            {

                                //add main life
                                Delay(1);
                                _driver.FindElement(By.XPath($"//*[@id='gatsby-focus-wrapper']/article/form/section[{section}]/div[2]/div[1]/div/label[{label}]")).Click();
                                //Cover Amount
                                DOB = _driver.FindElement(By.XPath($"/html/body/div[1]/div[1]/article/form/section[{section}]/div[3]/div[5]/input"));
                                date_of_birth = DOB.GetAttribute("value");
                                SlideBar(item["Cover_Amount"], lifeAsuredCounter, "Myself");
                                Delay(2);
                                 frontEndPrem = (_driver.FindElement(By.XPath($"/html/body/div[1]/div[1]/article/form/section[{section}]/div[4]/div[1]/label/h2/strong[2]")).Text).Remove(0,1).Trim();
                                frontEndMin = (_driver.FindElement(By.XPath($"/html/body/div[1]/div[1]/article/form/section[{section}]/div[4]/div[1]/div[2]/span[1]")).Text).Remove(0, 1).Replace(" ","");
                                frontEndMax = (_driver.FindElement(By.XPath($"/html/body/div[1]/div[1]/article/form/section[{section}]/div[4]/div[1]/div[2]/span[2]")).Text).Remove(0,1).Replace(" ", "");
                                validation = RolePlayerValidation(_driver, item["Cover_Amount"], "ML", date_of_birth, frontEndPrem , frontEndMin, frontEndMax);
                                if (validation.Item1 == "Failed")
                                {
                                    return Tuple.Create("Failed", validation.Item2);
                                }
                                //ID ,cover
                                section++;
                                lifeAsuredCounter++;
                                break;
                            }
                        }
                        //click Add 
                        Delay(2);
                        _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/button")).Click();

                        //select relationship
                        Delay(2);
                        _driver.FindElement(By.XPath($"//*[@id='gatsby-focus-wrapper']/article/form/section[{section}]/div[2]/div[1]/div/label[{label}]")).Click();
                        if (key == "Extended")
                        {
                            //Extended Relationship Type

                            IWebElement RelationshipType = _driver.FindElement(By.Name($"/cover-details[{lifeAsuredCounter}].relationship-extended-type"));
                            RelationshipType.SendKeys(item["Extended_RelationshipType"]);
                            RelationshipType.SendKeys(Keys.ArrowDown);
                            RelationshipType.SendKeys(Keys.Enter);
                        }
                        //FirstName
                        Delay(1);
                        _driver.FindElement(By.Name($"/cover-details[{lifeAsuredCounter}].name")).SendKeys(item["First_name"]);
                        //Surname
                        Delay(2);
                        _driver.FindElement(By.Name($"/cover-details[{lifeAsuredCounter}].surname")).SendKeys(item["Surname"]);
                        //ID Number
                        Delay(1);
                        _driver.FindElement(By.Name($"/cover-details[{lifeAsuredCounter}].id-number")).SendKeys(item["ID_number"]);
                        //MaxMin Age validation
                        Tuple<string, string> ageValidationResults = MaxMinAgeValidation(section);
                        if(ageValidationResults.Item1 != "" && ageValidationResults.Item2 !="")
                        {
                            return  Tuple.Create(ageValidationResults.Item1,ageValidationResults.Item2);
                        }
                        
           

                        //Cellphone
                        Delay(2);
                        _driver.FindElement(By.Name($"/cover-details[{lifeAsuredCounter}].contact-number")).SendKeys(item["Cellphone"]);
                        DOB = _driver.FindElement(By.XPath($"/html/body/div[1]/div[1]/article/form/section[{section}]/div[3]/div[5]/input"));
                        date_of_birth = DOB.GetAttribute("value");
                        SlideBar( item["Cover_Amount"], lifeAsuredCounter, key);
                        //
                        Delay(2);
                        frontEndPrem = (_driver.FindElement(By.XPath($"/html/body/div[1]/div[1]/article/form/section[{section}]/div[4]/div[1]/label/h2/strong[2]")).Text).Remove(0, 1).Trim();
                        frontEndMin = (_driver.FindElement(By.XPath($"/html/body/div[1]/div[1]/article/form/section[{section}]/div[4]/div[1]/div[2]/span[1]")).Text).Remove(0, 1).Replace(" ", "");
                        frontEndMax = (_driver.FindElement(By.XPath($"/html/body/div[1]/div[1]/article/form/section[{section}]/div[4]/div[1]/div[2]/span[2]")).Text).Remove(0, 1).Replace(" ", "");
                        validation = RolePlayerValidation(_driver, item["Cover_Amount"], key, date_of_birth, frontEndPrem, frontEndMin, frontEndMax);
                        if (validation.Item1 == "Failed")
                        {
                            return Tuple.Create("Failed", validation.Item2);
                        }



                        section++;
                        lifeAsuredCounter++;
                    }
                    else
                    {
                        break;
                    }
                }

                label++;

            }


            Delay(1);

            //Click next
            Delay(3);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div[1]/a[2]")).Click();


            var beneCounter = 0;
            //payment reciever(Beneficiary)
            foreach (var item in beneficiaries)
            {
                //click relationship 
                Delay(1);
                _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/div/section/div[3]/div/div[1]/div/label[1]")).Click();

                //FirstName
                Delay(1);
                _driver.FindElement(By.Name($"/funeral-beneficiaries[{beneCounter}].name")).SendKeys(item["First_name"]);
                //Surname
                Delay(1);
                _driver.FindElement(By.Name($"/funeral-beneficiaries[{beneCounter}].surname")).SendKeys(item["Surname"]);

                //ID Number
                Delay(1);
                _driver.FindElement(By.Name($"/funeral-beneficiaries[{beneCounter}].id-number")).SendKeys(item["ID_number"]);

                //Cellphone
                Delay(1);
                _driver.FindElement(By.Name($"/funeral-beneficiaries[{beneCounter}].contact-number")).SendKeys(item["Cellphone"]);

                //Percentage
                IWebElement sliderbar5 = _driver.FindElement(By.ClassName("slider"));
                int widthslider5 = sliderbar5.Size.Width;
                Delay(1);
                IWebElement slider5 = _driver.FindElement(By.ClassName("slider"));
                Actions slideraction5 = new Actions(_driver);
                slideraction5.ClickAndHold(slider5);
                slideraction5.MoveByOffset(260, 0).Build().Perform();
                beneCounter++;

            }


            //Click next
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div[1]/a[2]")).Click();




            //Word of advice
            Delay(1);
            _driver.FindElement(By.XPath("/html/body/div[1]/div[1]/article/form/section/div[2]/div/textarea")).SendKeys("Test");


            //Click next
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div[1]/a[2]")).Click();

            //Click No
            Delay(3);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/section/div[2]/form/div[2]/div/label[2]")).Click();

            //*[@id="gatsby-focus-wrapper"]/article/section/div[2]/form/div[2]/div/label[2]
            //go to payment 
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div/a[2]")).Click();

            /////////Payment Details
            string bank = policyHolderData["Bank"];

            //policy payer
            Delay(1);
            _driver.FindElement(By.Name("/same-as-fna")).Click();


            //bank details
            //Bank Selction

            SelectElement oSelect1 = new SelectElement(_driver.FindElement(By.Name("/bank-select")));
            oSelect1.SelectByValue(bank);

            //Account Number
            Delay(1);
            _driver.FindElement(By.Name("/account-number")).SendKeys(policyHolderData["Account_Number"]);


            //Account Type
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section[1]/div[2]/div[4]/div/label[2]")).Click();


            ///debit - order - date / debit - order - date
            SelectElement oSelect = new SelectElement(_driver.FindElement(By.Name("/debit-order-date")));
            oSelect.SelectByValue(policyHolderData["Debit_Order_Day"]);

            //salarypaydate
            Delay(1);
            _driver.FindElement(By.XPath("/html/body/div[1]/div[1]/article/form/section[1]/div[2]/div[6]/input")).SendKeys(policyHolderData["Salary_Date"]);

            //click tickbox
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='/arrange-payment-gather-information-disclaimer']")).Click();

            //click yes
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section[2]/section/div[1]/div/div/label[1]")).Click();

            //click yes
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section[2]/section/div[2]/div/div/label[1]")).Click();


            //click next
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div/a")).Click();


            //click i uderstand
            Delay(1);
            IWebElement iagree = _driver.FindElement(By.XPath("/html/body/reach-portal/div/div/div/button"));
            iagree.Click();

            //click start
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/section/div[3]/button")).Click();

            //debicheck loading delay
            //Impletent implicit wait
            WebDriverWait wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(160));
            try {
                
                wait.Until(ExpectedConditions.ElementExists(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div/a[2]")));
            }
            catch
            {
                var tries = 2;
                for(int i = 0; i < tries; i++)
                {
                    WebDriverWait wt = new WebDriverWait(_driver, TimeSpan.FromSeconds(160));
                    wt.Until(ExpectedConditions.ElementExists(By.XPath("/html/body/div[1]/div[1]/article/section/div[3]/button")));
                    _driver.FindElement(By.XPath("/html/body/div[1]/div[1]/article/section/div[3]/button")).Click();
                }
             
                
            }
           
           


            var Errormessage = _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/section/div[2]/div[2]/div[2]/p")).Text;

            if (Errormessage == "DebiCheck accepted by customer")
            {



                Delay(1);
                //Click next
                _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div/a[2]")).Click();



            }
            else
            {

                comment = "Debicheck Failed";
                results = "Failed";
     
                return Tuple.Create(results, comment);


            }


            //Physical Address

            //Building
            Delay(3);
            _driver.FindElement(By.Name("/physical-address-building")).SendKeys(policyHolderData["Building"]);
            //Street
            Delay(1);
            _driver.FindElement(By.Name("/physical-address-street")).SendKeys(policyHolderData["Street"]);

            //Town
            Delay(1);
            _driver.FindElement(By.Name("/physical-address-town")).SendKeys(policyHolderData["City"]);

            //Suburb
            Delay(1);
            _driver.FindElement(By.Name("/physical-address-suburb")).SendKeys(policyHolderData["Suburb"]);

            //CodeField 
            _driver.FindElement(By.Name("/physical-address-code")).SendKeys(policyHolderData["Code"]);



            ///click tickbox same-as-physical
            Delay(1);
            _driver.FindElement(By.Name("/same-as-physical")).Click();

            ///click tickbox 
            Delay(1);
            _driver.FindElement(By.Name("/policy-holder-signature-datetime")).Click();

            ///click tickbox 
            Delay(1);
            _driver.FindElement(By.Name("/premium-payer-signature-datetime")).Click();

            //reference no 
            Delay(1);
            _driver.FindElement(By.Name("/call-reference-number")).SendKeys(policyHolderData["Call_Ref_Number"]);

            //click next 
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div[1]/a[2]")).Click();


            ///click tickbox 
            Delay(1);
            _driver.FindElement(By.Name("/popia-consent-datetime")).Click();


            //click next 
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div/a[2]")).Click();



            Delay(2);
            //upload1
            _driver.FindElement(By.Id("/identification")).SendKeys(upload_file);

            //upload2
            Delay(2);
            _driver.FindElement(By.Id("/q-link")).SendKeys(upload_file);
            //upload3
            Delay(1);
            _driver.FindElement(By.Id("/proof-of-income")).SendKeys(upload_file);


            //click next
            Delay(8);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div/a[2]")).Click();

            var cardNum = "";
            for (int i = 0; i < 8; i++)
            {
               Random rnd = new Random();
                cardNum = cardNum + rnd.Next(10);
            }

            //Card number
            Delay(4);
            _driver.FindElement(By.Id("/card-number")).SendKeys(cardNum);

            //next
            Delay(2);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div[1]/a[2]")).Click();



            //next
            Delay(4);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/div[2]/div/a")).Click();




            //sync
            Delay(1);
            _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/nav/div/div/button")).Click();

            var appStatus = _driver.FindElement(By.CssSelector("#gatsby-focus-wrapper > article > div.card.tab-container > div.tab-body > section > section:nth-child(1) > div > div.final-block > span")).Text;
            

            for (int i = 0;i < 5;i++)
            {
            
                _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/nav/div/div/button")).Click();
                Delay(10);
                
                appStatus = _driver.FindElement(By.CssSelector("#gatsby-focus-wrapper > article > div.card.tab-container > div.tab-body > section > section:nth-child(1) > div > div.final-block > span")).Text;
                if(appStatus == "Uploaded")
                {
                    break;
                }
            }

            if(appStatus == "Uploaded")
            {
                results = "Passed";
            }
            else
            {
                results = "Failed";
                comment = "Application was not succesfull";
            }




            return Tuple.Create(results, comment);
        }

 

        private Tuple<string,string> MaxMinAgeValidation(int section)
        {
           
            try
            {
                string results = "";
                String validationMsg = _driver.FindElement(By.XPath($"//*[@id='gatsby-focus-wrapper']/article/form/section[{section}]/div[4]/div[1]/label")).Text;

                switch (validationMsg)
                {
                    case "Cover is only available for parents from 26 to 85 years of age":
                        TakeScreenshot(_driver, $@"{_screenShotFolder}\Failed_Scenarios\", "ParentAgeValidation");
                        results = "Failed";
                        return Tuple.Create(results, validationMsg);

                    case "Cover is only available for spouses from 18 to 64 years of age":
                        TakeScreenshot(_driver, $@"{_screenShotFolder}\Failed_Scenarios\", "SpouseAgeValidation");
                        results = "Failed";
                        return Tuple.Create(results, validationMsg);


                    case "Cover is only available for persons up to 85 years of age":
                        TakeScreenshot(_driver, $@"{_screenShotFolder}\Failed_Scenarios\", "ExtendedAgeValidation");
                        results = "Failed";
                        return Tuple.Create(results, validationMsg);
                    case "Cover is only available for children up to 25 years of age":
                        TakeScreenshot(_driver, $@"{_screenShotFolder}\Failed_Scenarios\", "ChildAgeValidation");
                        results = "Failed";
                        return Tuple.Create(results, validationMsg);
                }
            }
            catch
            {

                return Tuple.Create("", "");

            }

            return Tuple.Create("", "");

        }

        public void SlideBar(string coverAmount, int counts,string role)
        {
            
           
            

            if (role == "Myself")
            {

                var V_Position = "";
                switch (coverAmount)
                {

                    case "5000":
                        V_Position = "-500";
                        break;
                    case "7500":
                        V_Position = "-400";
                        break;
                    case "10000":
                        V_Position = "-300";
                        break;
                    case "15000":
                        V_Position = "-200";
                        break;
                    case "20000":
                        V_Position = "50";
                        break;

                    case "30000":
                        V_Position = "200";
                        break;
                    case "40000":
                        V_Position = "300";
                        break;
                    case "50000":
                        V_Position = "400";
                        break;
                    case "60000":
                        V_Position = "500";
                        break;

                }
                IWebElement sliderbar = _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section[3]/div[4]/div[1]"));

                int widthslider = sliderbar.Size.Width;
                Delay(1);
                IWebElement slider = _driver.FindElement(By.XPath($"//*[@id='/cover-details[{counts}].cover-amount']"));

                Actions slideraction = new Actions(_driver);
                slideraction.ClickAndHold(slider);
                Delay(1);
                //f = Mathf.Round(f * 100.0f) * 0.01f;
                slideraction.MoveByOffset(Convert.ToInt32(V_Position), 0).Build().Perform();

            }


            if (role == "Children" || role == "spouse")
            {

                
                    var V1_Position = "";
                    switch (coverAmount)
                    {

                        case "5000":
                            V1_Position = "-500";
                            break;
                        case "7500":
                            V1_Position = "-400";
                            break;
                        case "10000":
                            V1_Position = "-300";
                            break;
                        case "15000":
                            V1_Position = "-200";
                            break;
                        case "20000":
                            V1_Position = "50";
                            break;

                        case "30000":
                            V1_Position = "200";
                            break;
                        case "40000":
                            V1_Position = "300";
                            break;
                        case "50000":
                            V1_Position = "400";
                            break;
                        case "60000":
                            V1_Position = "500";
                            break;
                    }
                    IWebElement sliderbar = _driver.FindElement(By.XPath("//*[@id='gatsby-focus-wrapper']/article/form/section[3]/div[4]/div[1]"));

                    int widthslider = sliderbar.Size.Width;
                    Delay(1);
                    IWebElement slider = _driver.FindElement(By.XPath($"//*[@id='/cover-details[{counts}].cover-amount']"));
                    Actions slideraction = new Actions(_driver);
                    slideraction.ClickAndHold(slider);
                    Delay(1);
                    //f = Mathf.Round(f * 100.0f) * 0.01f;
                    slideraction.MoveByOffset(Convert.ToInt32(V1_Position), 0).Build().Perform();

                    
              


            }
            if (role == "Parents")
            {
              
                    var V2_Position = "";
                    switch (coverAmount)
                    {

                        case "5000":
                            V2_Position = "-500";
                            break;
                        case "7500":
                            V2_Position = "-400";
                            break;
                        case "10000":
                            V2_Position = "-300";
                            break;
                        case "15000":
                            V2_Position = "-200";
                            break;
                        case "20000":
                            V2_Position = "50";
                            break;

                    }

                    IWebElement sliderbar = _driver.FindElement(By.ClassName("slider"));
                    int widthslider = sliderbar.Size.Width;
                    Delay(1);
                    IWebElement slider = _driver.FindElement(By.XPath($"//*[@id='/cover-details[{counts}].cover-amount']"));
                    Actions slideraction = new Actions(_driver);
                    slideraction.ClickAndHold(slider);
                    slideraction.MoveByOffset(Convert.ToInt32(V2_Position), 0).Build().Perform();



                

                if (role == "Extended")
                {
                   
                        var V3_Position = "";
                        switch (coverAmount)
                        {

                            case "5000":
                                V3_Position = "-500";
                                break;
                            case "7500":
                                V3_Position = "-400";
                                break;
                            case "10000":
                                V3_Position = "-300";
                                break;
                            case "15000":
                                V3_Position = "-200";
                                break;
                            case "20000":
                                V3_Position = "50";
                                break;
                            case "30000":
                                V3_Position = "200";
                                break;
                        }
                        IWebElement sliderbar2 = _driver.FindElement(By.ClassName("slider"));
                        int widthslider2 = sliderbar.Size.Width;
                        Delay(1);
                        IWebElement slider2 = _driver.FindElement(By.XPath($"//*[@id='/cover-details[{counts}].cover-amount']"));
                        Actions slideraction2 = new Actions(_driver);
                        slideraction.ClickAndHold(slider);
                        slideraction.MoveByOffset(Convert.ToInt32(V3_Position), 0).Build().Perform();




                }



            }
         



}


        public Tuple<string, string> RolePlayerValidation(IWebDriver _driver ,string coverAmount,string roleplayer,string dob, string expectedPrem, string frondEndMin, string frondEndMax)
        {

            //calulate age
            String premValidation = "", coverAmountsValidation = "", comment = "", age;
            var birthYear = dob.Split("-")[0];
            var birthMonth = dob.Split("-")[1];
            var birthDay = dob.Split("-")[2];

            age = (DateTime.Now.Year - Convert.ToInt32(birthYear)).ToString();

            if (Convert.ToInt32(birthMonth) > DateTime.Now.Month || (Convert.ToInt32(birthMonth) == DateTime.Now.Month && Convert.ToInt32(birthDay) > DateTime.Now.Day ))
            {
                age = (Convert.ToInt32(age) - 1).ToString();
            }
           


            Delay(2);
            var premiumFromRate = getPremuimFromRateTable(age, "ML", coverAmount, "Serenity_Funeral_Core");


            //Validate Cover Limits
            using (OleDbConnection conn = new OleDbConnection(_test_data_connString))
            {
                try
                {
                    var sheet = "Limits";
                    // Open connection
                    conn.Open();
                    string cmdQuery = $"SELECT * FROM [{sheet}$]";

                    OleDbCommand cmd = new OleDbCommand(cmdQuery, conn);

                    // Create new OleDbDataAdapter
                    OleDbDataAdapter oleda = new OleDbDataAdapter();

                    oleda.SelectCommand = cmd;

                    // Create a DataSet which will hold the data extracted from the worksheet.
                    DataSet ds = new DataSet();

                    // Fill the DataSet from the data extracted from the worksheet.
                    oleda.Fill(ds, "Policies");


                    //addMainLife();
                    foreach (var row in ds.Tables[0].DefaultView)
                    {
                        //""
                        var rolePl = ((System.Data.DataRowView)row).Row.ItemArray[4].ToString();
                        var product = ((System.Data.DataRowView)row).Row.ItemArray[5].ToString();

                        if (rolePl == roleplayer && product == "Serenity_Funeral_Core")
                        {
                            var minAge = ((System.Data.DataRowView)row).Row.ItemArray[0].ToString();
                            var maxAge = ((System.Data.DataRowView)row).Row.ItemArray[1].ToString();
                            var minCover = ((System.Data.DataRowView)row).Row.ItemArray[2].ToString();
                            var maxCover = ((System.Data.DataRowView)row).Row.ItemArray[3].ToString();
                            //Check if age falls between ages from spreadsheet
                            if (Convert.ToInt32(age) >= Convert.ToInt32(minAge) && Convert.ToInt32(age) <= Convert.ToInt32(maxAge))
                            {
                                //check the amount is between 
                                if(minCover == frondEndMin && maxCover == frondEndMax)
                                {
                                    coverAmountsValidation = "Passed";
                                }
                                else
                                {
                                    coverAmountsValidation = "Failed";
                                    comment += $"Scenario Failed - max and min cover amounts validation erorr for {roleplayer} age {age} ";
                                }

                            }
                       
                           
                        }

                    }
                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                }
            }




            if (premiumFromRate != Convert.ToDecimal(expectedPrem))
            {
                premValidation = "Failed";
                comment += $"Scenario Failed premium validation for {roleplayer}. Premuim on frontend does not match one in rate table";

            }
            else
            {
                premValidation = "Passed";
            }

           if(premValidation == "Passed" && coverAmountsValidation == "Passed")
           {
                return Tuple.Create("Passed", "");
           }
           else
           {
                
                TakeScreenshot(_driver, $"C:/Users/G992127/Documents/GitHub/ILR_TestSuite/ILR_TestSuite/New Business/{DateTime.Now.Year + DateTime.Now.Month + DateTime.Now.Day}​/validations/", $"{roleplayer}_premiumValidation");
                return Tuple.Create("Failed", comment);


           }



        }

            [TearDown]
        public void closeBrowser()
        {
            base.DisconnectBrowser();
        }

    }
}
