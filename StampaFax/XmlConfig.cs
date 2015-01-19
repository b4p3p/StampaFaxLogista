using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.IO;
using System.Globalization;

namespace StampaFax
{
    public class XmlConfig
    {
        XmlDocument xmlConfig;

        public XmlConfig()
        {
            try
            {
                xmlConfig = new XmlDocument();
                xmlConfig.Load("Config.xml");

                string app;
                app = xmlConfig.SelectSingleNode("//giornoprelievo").InnerText;
                app = xmlConfig.SelectSingleNode("//ultimoordine").InnerText;


            }
            catch (Exception)
            {
                xmlConfig.LoadXml(@"
                <config>
                    <giornoprelievo>1</giornoprelievo>
                    <ultimoordine>1</ultimoordine>
                </config>
                ");

                if (File.Exists("Config.xml")) File.Delete("Config.xml");

                xmlConfig.Save("Config.xml");
            }
        }

        public int UltimoOrdine
        {
            get
            {
                try
                {
                    int ris = 0;
                    int.TryParse(xmlConfig.SelectSingleNode("//ultimoordine").InnerText, out ris);
                    return ris;
                }
                catch (Exception)
                {
                    return 1;
                }
            }

            set
            {
                try
                {
                    xmlConfig.SelectSingleNode("//ultimoordine").InnerText = value.ToString();
                }
                catch (Exception)
                {
                    xmlConfig.SelectSingleNode("//ultimoordine").InnerText = "1";
                }
                Salva();
            }
        }

        public int GiornoSettimanaPrelievo
        {
            get
            {
                try
                {
                    int ris = 0;
                    int.TryParse(xmlConfig.SelectSingleNode("//giornoprelievo").InnerText, out ris);
                    return ris;

                }
                catch (Exception)
                {
                    return 0;
                }                
            }

            set
            {
                try
                {
                    xmlConfig.SelectSingleNode("//giornoprelievo").InnerText = value.ToString();
                }
                catch (Exception)
                {
                    xmlConfig.SelectSingleNode("//giornoprelievo").InnerText = "1";
                }

                Salva();
            }
        }

        public void Salva()
        {
            xmlConfig.Save("Config.xml");
        }



        internal String GetProssimoDataPrelievo()
        {
            DateTime risData = DateTime.Now;
            
            int giornoSettimana = (int) risData.DayOfWeek ;

            for (int i = 0; i < 7; i++)
            {
                risData = risData.AddDays(1);
                if ((int)risData.DayOfWeek == GiornoSettimanaPrelievo) break;
            }

            return risData.Day.ToString() + "/" + risData.Month.ToString() + "/" + risData.Year.ToString(); 
        }
    }
}
