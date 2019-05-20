using System;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using System.Collections.Generic;
using OpenQA.Selenium.Chrome;


namespace WindowsAppStoreScrapper
{
    public class Program
    {
        public static IWebDriver driver = new ChromeDriver();
        public static List<IWebElement> categories = new List<IWebElement>();
        public static string baseURL = "https://www.microsoft.com";

        public static void TheWebScrapperTest()
        {
            driver.Navigate().GoToUrl(baseURL + "/en-us/store/apps/windows-phone?icid=CNavAppsWindowsPhoneApps"); //visiting the app store
            //retrieving the catagories holder
            driver.FindElement(By.Id("R1MarketRedirect-close")).Click();
            IWebElement catagoryHolder = driver.FindElement(By.CssSelector(".m-multi-column.f-columns-4"));
            categories.AddRange(catagoryHolder.FindElements(By.TagName("a")));

            for (int j = 0; j < categories.Count; j++)
            {
                //Clicking each catagories
                string categoryName = categories[j].Text;
                categories[j].Click();
                //finding apps
                List<IWebElement> apps = new List<IWebElement>();
                apps.AddRange(driver.FindElements(By.CssSelector("h3.c-heading")));
                driver.FindElement(By.CssSelector(".c-text-field.f-flex.newsletter-email-field")).SendKeys("abc@abc.com");
                driver.FindElement(By.CssSelector(".ctaEmailNewsletterSignUpButton.c-button")).Click();
                getNextPagesApps:
                for (int i = 0; i < apps.Count; i++)
                {
                    apps[i].Click();
                    //Checking which version

                    try
                    {
                        Thread.Sleep(5000);
                        if (driver.FindElement(By.Id("page-title")).Displayed)
                        {
                            string appTitle = driver.FindElement(By.Id("page-title")).Text;
                            IWebElement productDetails = driver.FindElement(By.ClassName("context-product-details"));

                            string publisher = productDetails.FindElement(By.TagName("div")).Text;

                            IWebElement priceHolder = driver.FindElement(By.ClassName("price-info"));
                            string price = priceHolder.FindElement(By.TagName("span")).Text;

                            //extracting the review part
                            IWebElement ratingHolder = driver.FindElement(By.ClassName("m-histogram"));
                            string rating = ratingHolder.FindElement(By.TagName("span")).Text;

                            int totalComments = Int32.Parse(driver.FindElement(By.ClassName("c-meta-text")).Text);
                            collectReviewsAgain:
                            Thread.Sleep(5000);
                            List<IWebElement> reviews = new List<IWebElement>();
                            reviews.AddRange(driver.FindElements(By.CssSelector(".review.cli_review")));

                            foreach (IWebElement review in reviews)
                            {


                                double commentRating = Double.Parse(review.FindElement(By.CssSelector(".c-rating.f-user-rated.f-individual")).GetAttribute("data-value"));
                                string commentDate = review.FindElements(By.ClassName("c-paragraph-3"))[0].Text;
                                string reviewerName = review.FindElements(By.ClassName("c-paragraph-3"))[1].Text;

                                string platform = "NA";
                                try { platform = review.FindElement(By.ClassName("c-meta-text")).Text; } catch (Exception c) { };

                                string titleComment = "NA";
                                try { titleComment = review.FindElement(By.ClassName("c-heading-6")).Text; } catch (Exception c) { };

                                //extracting total comment
                                string fullComment = "NA";
                                try
                                {
                                    IWebElement fullCommentHolder = review.FindElement(By.CssSelector(".c-content-toggle.cli_reviews_content_toggle"));
                                    fullComment = fullCommentHolder.FindElement(By.TagName("p")).Text;
                                }
                                catch (Exception c) { };

                                string usefulness = "NA";
                                try { usefulness = review.FindElements(By.ClassName("c-meta-text"))[1].Text; } catch (Exception c) { };

                                //Console.WriteLine("Category: " + categoryName + "App: " + appTitle + " reviewer: " + reviewerName + " ");
                                using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"E:\Kids and Family.csv", true))
                                {
                                    file.Write(categoryName + "," + appTitle + "," + price + "," + publisher + "," + rating + "," + commentRating + "," + commentDate + "," + reviewerName + "," + platform + "," + titleComment + "," + fullComment + "," + usefulness + "\n");
                                }

                            }
                            try
                            {
                                if (driver.FindElement(By.Id("reviewsPageNextAnchor")).Displayed)
                                {
                                    driver.FindElement(By.Id("reviewsPageNextAnchor")).Click();
                                    goto collectReviewsAgain;
                                }
                                else
                                {
                                    driver.Navigate().Back();

                                }
                            }
                            catch (Exception e)
                            {
                                driver.Navigate().Back();

                            }
                        }
                    }
                    catch (Exception e)
                    {
                        IWebElement appTitleHolder = driver.FindElement(By.Id("productTitle"));
                        string appTitle = appTitleHolder.FindElement(By.TagName("h1")).Text;

                        IWebElement productDetails = driver.FindElement(By.ClassName("buybox-metadata "));

                        string publisher = productDetails.FindElement(By.TagName("span")).Text;
                        string price = driver.FindElement(By.ClassName("pi-price-text")).Text;

                        //extracting the review part

                        driver.FindElement(By.Id("pivot-tab-ReviewsTab")).Click();

                        collectReviewsAgain:
                        Thread.Sleep(5000);
                        List<IWebElement> reviews = new List<IWebElement>();
                        reviews.AddRange(driver.FindElements(By.CssSelector(".review.cli_review")));

                        foreach (IWebElement review in reviews)
                        {
                            IWebElement ratingHolder = driver.FindElement(By.ClassName("m-histogram"));
                            string rating = ratingHolder.FindElement(By.TagName("span")).Text;

                            double commentRating = Double.Parse(review.FindElement(By.CssSelector(".c-rating.f-user-rated.f-individual")).GetAttribute("data-value"));
                            List<IWebElement> combinedAuthorAndDate = new List<IWebElement>();
                            combinedAuthorAndDate.AddRange(review.FindElements(By.ClassName("c-paragraph-3")));
                            string commentDate = combinedAuthorAndDate[0].Text;
                            string reviewerName = combinedAuthorAndDate[1].Text;
                            string platform = "NA";
                            try { platform = review.FindElement(By.ClassName("c-meta-text")).Text; } catch (Exception c) { };
                            string titleComment = "NA";
                            try { titleComment = review.FindElement(By.ClassName("c-heading-6")).Text; } catch (Exception c) { };

                            //extracting total comment
                            string fullComment = "NA";
                            try
                            {
                                IWebElement fullCommentHolder = review.FindElement(By.CssSelector(".c-content-toggle.cli_reviews_content_toggle"));
                                fullComment = fullCommentHolder.FindElement(By.TagName("p")).Text;
                            }
                            catch (Exception c) { };
                            string usefulness = "NA";
                            try { usefulness = review.FindElements(By.ClassName("c-meta-text"))[1].Text; } catch (Exception c) { };

                            //Console.WriteLine("Category: " + categoryName + "App: " + appTitle + " reviewer: " + reviewerName + " ");
                            //writing to file
                            using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"D:\AppAnalytics.csv", true))
                            {
                                file.Write(categoryName + "," + appTitle + "," + price + "," + publisher + "," + rating + "," + commentRating + "," + commentDate + "," + reviewerName + "," + platform + "," + titleComment + "," + fullComment + "," + usefulness + "\n");
                            }



                        }
                        //Checking next page's review
                        try
                        {
                            if (driver.FindElement(By.Id("reviewsPageNextAnchor")).Displayed)
                            {
                                driver.FindElement(By.Id("reviewsPageNextAnchor")).Click();
                                goto collectReviewsAgain;
                            }
                            else
                            {
                                driver.Navigate().Back();
                            }
                        }
                        catch (Exception ex)
                        {
                            driver.Navigate().Back();
                        }
                    }

                    apps.Clear();
                    apps.AddRange(driver.FindElements(By.CssSelector("h3.c-heading")));
                }

                if (driver.FindElement(By.PartialLinkText("Next")).Displayed)
                {
                    driver.FindElement(By.PartialLinkText("Next")).Click();
                    apps.Clear();
                    apps.AddRange(driver.FindElements(By.CssSelector("h3.c-heading")));
                    goto getNextPagesApps;

                }
                else
                {
                    driver.Navigate().GoToUrl("https://www.microsoft.com/en-us/store/apps/windows-phone?icid=CNavAppsWindowsPhoneApps");
                }
                categories.Clear();
                categories.AddRange(catagoryHolder.FindElements(By.TagName("a")));

            }
            //create new categories again



        }

        public static void nextPageApps()
        {
            if (driver.FindElement(By.PartialLinkText("Next")).Displayed)
            {
                driver.FindElement(By.PartialLinkText("Next")).Click();

            }
            else
            {
                driver.Navigate().Back();
            }
        }
        //private bool IsElementPresent(By by)
        //{
        //    try
        //    {
        //        driver.FindElement(by);
        //        return true;
        //    }
        //    catch (NoSuchElementException)
        //    {
        //        return false;
        //    }
        //}

        //private bool IsAlertPresent()
        //{
        //    try
        //    {
        //        driver.SwitchTo().Alert();
        //        return true;
        //    }
        //    catch (NoAlertPresentException)
        //    {
        //        return false;
        //    }
        //}

        //private string CloseAlertAndGetItsText()
        //{
        //    try
        //    {
        //        IAlert alert = driver.SwitchTo().Alert();
        //        string alertText = alert.Text;
        //        if (acceptNextAlert)
        //        {
        //            alert.Accept();
        //        }
        //        else
        //        {
        //            alert.Dismiss();
        //        }
        //        return alertText;
        //    }
        //    finally
        //    {
        //        acceptNextAlert = true;
        //    }
        //}
        public static void Main(string[] args)
        {
            Console.WriteLine("Inside main");
            TheWebScrapperTest();

        }
    }
}





