﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NormalizarCP.Entidades;
using NormalizarCP.Datos;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using System.Threading;


namespace NormalizarCP
{
    class Program
    {
        static bool success;

        static void Main(string[] args)
        {
            List<Calle> lista = new List<Calle>();

            lista = CalleDAO.readCalles(lista);

            IWebDriver driver = new FirefoxDriver();

            driver.Navigate().GoToUrl("https://codigopostal.com.ar/");

            foreach(Calle calle in lista)
            {
                success = false;
                while (!success)
                {
                    IWebElement textbox = driver.FindElement(By.XPath("//*[@id='address']"));
                    textbox.Clear();
                    textbox.SendKeys(calle.calle + " " + calle.altura_ini + " CABA");
                    driver.FindElement(By.XPath("//*[@id='submit']")).Click();
                    Thread.Sleep(2000);
                    try
                    {
                        calle.cp = driver.FindElement(By.XPath("//*[@id='results']/section/ul/li")).Text;
                        success = true;
                    }
                    catch (Exception)
                    {
                        driver.Navigate().Refresh();
                        success = false;
                        continue;
                    }

                    if (calle.cp.Contains("manual"))
                    {
                        calle.cp = "";
                    }
                    else
                    {
                        calle.cp = calle.cp.Substring(calle.cp.IndexOf(",") + 1);
                        calle.cp = calle.cp.Substring(0, calle.cp.Length - calle.cp.IndexOf(",") + 3);
                        calle.cp = calle.cp.Replace("CABA", "").Trim();
                    }

                    CalleDAO.cpToExcel(calle, "codigosPostales.xlsx");
                }
                
            }

            driver.Quit();

            
        }
    }
}
