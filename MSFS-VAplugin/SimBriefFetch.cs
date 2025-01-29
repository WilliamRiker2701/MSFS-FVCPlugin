using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using System.Globalization;
using System.Net.Http;
using System.Xml;
using System.IO;
using FSUIPC;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace MSFS
{
    public class SimBriefFetch
    {

        public string Flight { get; set; } = "";
        public string AirlineICAO { get; set; } = "";
        public string CallSign { get; set; } = "";
        public string CruiseProf { get; set; }
        public string ClimbProf { get; set; }
        public string DescentProf { get; set; }
        public int CostIndex { get; set; }
        public int InitialAlt { get; set; }
        public int AvgWindComp { get; set; }
        public int AvgWindDir { get; set; }
        public int AvgWindSpd { get; set; }
        public double TopClimbOAT { get; set; }
        public string Route { get; set; }
        public string Origin { get; set; }
        public int OriginElevation { get; set; }
        public string OriginRwy { get; set; }
        public int OriginTransAlt { get; set; }
        public int OriginTransLevel { get; set; }
        public string OriginMetar { get; set; }
        public string OriginMetarTime { get; set; }
        public string OriginTAF { get; set; }
        public string OriginTAFTime { get; set; }
        public string OriginWind { get; set; }
        public string OriginPressure { get; set; }
        public double OriginPressureINHG {  get; set; }
        public string OriginTemp { get; set; }
        public string Destination { get; set; }
        public int DestElevation { get; set; }
        public string DestRwy { get; set; }
        public int DestTransAlt { get; set; }
        public int DestTransLevel { get; set; }
        public string DestMetar { get; set; }
        public string DestMetarTime { get; set; }
        public string DestTAF { get; set; }
        public string DestTAFTime { get; set; }
        public string DestWind { get; set; }
        public string DestPressure { get; set; }
        public double DestPressureINHG { get; set; }
        public string DestTemp { get; set; }
        public string Altn { get; set; }
        public int AltnElevation { get; set; }
        public string AltnRwy { get; set; }
        public int AltnTransAlt { get; set; }
        public int AltnTransLevel { get; set; }
        public string AltnMetar { get; set; }
        public string AltnMetarTime { get; set; }
        public string AltnTAF { get; set; }
        public string AltnTAFTime { get; set; }
        public string AltnWind { get; set; }
        public string AltnPressure { get; set; }
        public double AltnPressureINHG { get; set; }
        public string AltnTemp { get; set; }
        public string Units { get; set; }
        public double FinRes { get; set; }
        public double AltnFuel { get; set; }
        public double FinresPAltn { get; set; }
        public double Fuel { get; set; }
        public int Passenger { get; set; }
        public int Bags { get; set; }
        public double WeightPax { get; set; }
        public double WeightCargo { get; set; }
        public double Payload { get; set; }
        public double ZFW { get; set; }
        public double TOW { get; set; }

        


        protected async Task<string> GetHttpContent(HttpResponseMessage response)
        {
            return await response.Content.ReadAsStringAsync();
        }

        protected XmlNode FetchOnline()
        {


            HttpClient httpClient = new HttpClient();
            HttpResponseMessage response = httpClient.GetAsync(string.Format(Utils.simbriefURL, Utils.pilotID)).Result;

            VoiceAttackPlugin.LogOutput("Requesting...", "grey");

            if (response.IsSuccessStatusCode)
            {
                string responseBody = GetHttpContent(response).Result;
                if (responseBody != null && responseBody.Length > 0)
                {
                    VoiceAttackPlugin.LogOutput("HTTP Request succeded", "grey");
                    XmlDocument xmlDoc = new XmlDocument();
                    xmlDoc.LoadXml(responseBody);
                    return xmlDoc.ChildNodes[1];
                }
                else
                {
                    VoiceAttackPlugin.LogOutput("Response Body is empty", "grey");
                }
            }
            else
            {
                VoiceAttackPlugin.LogOutput("HTTP Request failed", "grey");
            }

            return null;
        }


        protected XmlNode LoadOFP()
        {
                    
            VoiceAttackPlugin.LogOutput("Fetching...", "grey");
            return FetchOnline();

        }

        public bool Load()
        {
            XmlNode sbOFP = LoadOFP();
            VoiceAttackPlugin.LogOutput("Importing data...", "grey");
            int count = 0;
            string[] windFixIdents = {"","", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
            int[] windFixAlts = { 0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0 };

            try
            {
                Flight = sbOFP["general"]["flight_number"].InnerText;
                VoiceAttackPlugin.LogOutput("Flight", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("Flight. " + e.Message, "red");
            }

            try
            {
                AirlineICAO = sbOFP["general"]["icao_airline"].InnerText;
                VoiceAttackPlugin.LogOutput("AirlineICAO", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("AirlineICAO. " + e.Message, "red");
            }


            try
            {
                CruiseProf = sbOFP["general"]["cruise_profile"].InnerText;
                VoiceAttackPlugin.LogOutput("CruiseProf", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("CruiseProf. " + e.Message, "red");
            }

            try
            {
                ClimbProf = sbOFP["general"]["climb_profile"].InnerText;
                VoiceAttackPlugin.LogOutput("ClimbProf", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("ClimbProf. " + e.Message, "red");
            }

            try
            {
                DescentProf = sbOFP["general"]["descent_profile"].InnerText;
                VoiceAttackPlugin.LogOutput("DescentProf", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("DescentProf. " + e.Message, "red");
            }

            try
            {
                CostIndex = Convert.ToInt32(sbOFP["general"]["costindex"].InnerText);
                VoiceAttackPlugin.LogOutput("CostIndex", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("CostIndex. " + e.Message, "red");
            }

            try
            {
                InitialAlt = Convert.ToInt32(sbOFP["general"]["initial_altitude"].InnerText);
                VoiceAttackPlugin.LogOutput("InitialAlt", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("InitialAlt. " + e.Message, "red");
            }

            try
            {
                AvgWindComp = Convert.ToInt32(sbOFP["general"]["avg_wind_comp"].InnerText);
                VoiceAttackPlugin.LogOutput("AvgWindComp", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("AvgWindComp. " + e.Message, "red");
            }

            try
            {
                AvgWindDir = Convert.ToInt32(sbOFP["general"]["avg_wind_dir"].InnerText);
                VoiceAttackPlugin.LogOutput("AvgWindDir", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("AvgWindDir. " + e.Message, "red");
            }

            try
            {
                AvgWindSpd = Convert.ToInt32(sbOFP["general"]["avg_wind_spd"].InnerText);
                VoiceAttackPlugin.LogOutput("AvgWindSpd", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("AvgWindSpd. " + e.Message, "red");
            }

            try
            {
                XmlNodeList fixNodes = sbOFP.SelectNodes("/OFP/navlog/fix");

                foreach (XmlNode fixNode in fixNodes)
                {
                    // Retrieve attributes from the current <fix> element
                    if (fixNode["ident"].InnerText == "TOC")
                    {
                        TopClimbOAT = Convert.ToDouble(fixNode["oat"].InnerText);

                    }


                }
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("FixNodes. " + e.Message, "red");
            }

            try
            {
                XmlNodeList fixNodes = sbOFP.SelectNodes("/OFP/navlog/fix");
                int i = 0;
                string fixName = "Fix";

                
                foreach (XmlNode fixNode in fixNodes)
                {
                    // Retrieve attributes from the current <fix> element
                    i++;
                    fixName = "Fix" + i.ToString();

                    VoiceAttackPlugin.LogMonitorOutput(fixName, "orange");
                    VoiceAttackPlugin.SetText("sb" + fixName + "ident", fixNode["ident"].InnerText);
                    VoiceAttackPlugin.SetText("sb" + fixName + "name", fixNode["name"].InnerText);
                    VoiceAttackPlugin.SetText("sb" + fixName + "type", fixNode["type"].InnerText);
                    VoiceAttackPlugin.SetText("sb" + fixName + "type", fixNode["type"].InnerText);
                    VoiceAttackPlugin.SetText("sb" + fixName + "freq", fixNode["frequency"].InnerText);
                    VoiceAttackPlugin.SetText("sb" + fixName + "via", fixNode["via_airway"].InnerText);
                    VoiceAttackPlugin.SetText("sb" + fixName + "isSidStar", fixNode["is_sid_star"].InnerText);
                    VoiceAttackPlugin.SetText("sb" + fixName + "alt", fixNode["altitude_feet"].InnerText);
                    VoiceAttackPlugin.SetText("sb" + fixName + "spd", fixNode["ind_airspeed"].InnerText);
                    VoiceAttackPlugin.SetText("sb" + fixName + "windSpd", fixNode["wind_spd"].InnerText);
                    VoiceAttackPlugin.SetText("sb" + fixName + "windDir", fixNode["wind_dir"].InnerText);
                    VoiceAttackPlugin.SetText("sbFixNumber", i.ToString());

                    if (fixNode["ident"].InnerText == "TOD")
                    {
                        VoiceAttackPlugin.SetText("sbFirstDESCFix", i.ToString());                        
                    }


                    if (fixNode["ident"].InnerText == "TOC")
                    {
                        VoiceAttackPlugin.SetText("sbFirstCRZFix", i.ToString());
                    }

                }
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("FixNodes. " + e.Message, "red");
            }                      

            try
            {
                Route = sbOFP["general"]["route"].InnerText;
                VoiceAttackPlugin.LogOutput("Route", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("Route. " + e.Message, "red");
            }

            try
            {
                Origin = sbOFP["origin"]["icao_code"].InnerText;
                VoiceAttackPlugin.LogOutput("Origin", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("Origin. " + e.Message, "red");
            }

            try
            {
                OriginElevation = Convert.ToInt32(sbOFP["origin"]["elevation"].InnerText);
                VoiceAttackPlugin.LogOutput("OriginElevation", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("OriginElevation. " + e.Message, "red");
            }

            try
            {
                OriginRwy = sbOFP["origin"]["plan_rwy"].InnerText;
                VoiceAttackPlugin.LogOutput("OriginRwy", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("OriginRwy. " + e.Message, "red");
            }

            try
            {
                OriginTransAlt = Convert.ToInt32(sbOFP["origin"]["trans_alt"].InnerText);
                VoiceAttackPlugin.LogOutput("OriginTransAlt", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("OriginTransAlt. " + e.Message, "red");
            }

            try
            {
                OriginTransLevel = Convert.ToInt32(sbOFP["origin"]["trans_level"].InnerText);
                VoiceAttackPlugin.LogOutput("OriginTransLevel", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("OriginTransLevel. " + e.Message, "red");
            }

            try {
                OriginMetar = sbOFP["origin"]["metar"].InnerText;
                VoiceAttackPlugin.LogOutput("OriginMetar", "grey");
                }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("OriginMetar. " + e.Message, "red");
            }

            try
            {
                OriginMetarTime = sbOFP["origin"]["metar_time"].InnerText;
                VoiceAttackPlugin.LogOutput("OriginMetarTime", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("OriginMetarTime. " + e.Message, "red");
            }

            try
            {
                OriginTAF = sbOFP["origin"]["taf"].InnerText;
                VoiceAttackPlugin.LogOutput("OriginTAF", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("OriginTAF. " + e.Message, "red");
            }

            try
            {
                OriginTAFTime = sbOFP["origin"]["taf_time"].InnerText;
                VoiceAttackPlugin.LogOutput("OriginTAFTime", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("OriginTAFTime. " + e.Message, "red");
            }

            try
            {
                if (OriginTAFTime.Length > 1 && OriginTAF.IndexOf("KT ") != -1)
                {
                    Utils.sbOriginAvailWeather = 2;
                    Utils.sbWeatherReportCharPos = OriginTAF.IndexOf("KT ");
                    while (OriginTAF.Substring(Utils.sbWeatherReportCharPos,1) != " ")
                    {
                        Utils.sbWeatherReportCharPos--;
                    }
                    Utils.sbWeatherReportCharPos++;

                    OriginWind = OriginTAF.Substring(Utils.sbWeatherReportCharPos, OriginTAF.IndexOf("KT ")+2-Utils.sbWeatherReportCharPos);
                    VoiceAttackPlugin.LogOutput("OriginWind", "grey");

                }
                else if (OriginMetarTime.Length > 1 && OriginMetar.IndexOf("KT ") != -1)
                {
                    Utils.sbOriginAvailWeather = 1;
                    Utils.sbWeatherReportCharPos = OriginMetar.IndexOf("KT ");
                    while (OriginMetar.Substring(Utils.sbWeatherReportCharPos,1) != " ")
                    {
                        Utils.sbWeatherReportCharPos--;
                    }
                    Utils.sbWeatherReportCharPos++;

                    OriginWind = OriginMetar.Substring(Utils.sbWeatherReportCharPos, OriginMetar.IndexOf("KT ") + 2 - Utils.sbWeatherReportCharPos);
                    VoiceAttackPlugin.LogOutput("OriginWind", "grey");

                }
                else
                {
                    Utils.sbOriginAvailWeather = 0;
                    OriginWind = "NO DATA";
                }
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("OriginWind. " + e.Message, "red");
            }

            try
            {
                double presHPA = 0;
                double presINHG = 0;
                string presHPAs = "NULL";
                string presINHGs = "NULL";

                if (OriginMetarTime.Length > 1)
                {

                    if (OriginMetar.IndexOf(" Q") != -1)
                    {

                        Utils.sbWeatherReportCharPos = OriginMetar.IndexOf(" Q");
                        OriginPressure = OriginMetar.Substring(Utils.sbWeatherReportCharPos + 1, 5);

                    }

                    else if (OriginMetar.IndexOf(" A") != -1)
                    {

                        Utils.sbWeatherReportCharPos = OriginMetar.IndexOf(" A");
                        OriginPressure = OriginMetar.Substring(Utils.sbWeatherReportCharPos + 1, 5);

                    }

                    else
                    {

                        OriginPressure = "NO DATA";

                    }

                    if (OriginPressure.Substring(0, 1) == "Q")
                    {
                        presHPAs = OriginPressure.Substring(1, 4);
                        presHPA = Convert.ToDouble(presHPAs);

                        presINHG = presHPA / 33.864;
                        OriginPressureINHG = presINHG;
                        presINHGs = Convert.ToString(presINHG);
                        presINHGs = presINHGs.Substring(0, 5);

                        OriginPressure = presHPAs + " hPa / " + presINHGs + " inHg";

                    }
                    else if (AltnPressure.Substring(0, 1) == "A")
                    {
                        presINHGs = OriginPressure.Substring(1, 4);
                        presINHG = Convert.ToDouble(presINHGs);
                        presINHG = presINHG / 100;
                        OriginPressureINHG = presINHG;

                        presHPA = presINHG * 33.864;
                        presHPAs = Convert.ToString(presHPA);
                        presHPAs = presHPAs.Substring(0, 4);

                        OriginPressure = presHPAs + " hPa / " + presINHGs + " inHg";
                    }

                    VoiceAttackPlugin.LogOutput("OriginPressure", "grey");

                }
                else
                {

                    OriginPressure = "NO DATA";

                }


            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("OriginPressure. " + e.Message, "red");
            }

            try
            {
                if (OriginMetarTime.Length > 1)
                {

                    if (OriginMetar.IndexOf("/") != -1 && (OriginMetar.Substring(OriginMetar.IndexOf("/") + 3, 1) == " " || OriginMetar.Substring(OriginMetar.IndexOf("/") + 4, 1) == " "))
                    {

                        Utils.sbWeatherReportCharPos = OriginMetar.IndexOf("/");
                        while (OriginMetar.Substring(Utils.sbWeatherReportCharPos, 1) != " ")
                        {
                            Utils.sbWeatherReportCharPos--;
                        }
                        Utils.sbWeatherReportCharPos++;
                        Utils.sbWeatherReportCharPos2 = Utils.sbWeatherReportCharPos;
                        while (OriginMetar.Substring(Utils.sbWeatherReportCharPos, 1) != " ")
                        {
                            Utils.sbWeatherReportCharPos++;
                        }
                        Utils.sbWeatherReportCharPos--;
                        OriginTemp = OriginMetar.Substring(Utils.sbWeatherReportCharPos2, Utils.sbWeatherReportCharPos+1 - Utils.sbWeatherReportCharPos2);

                    }

                    else
                    {

                        OriginTemp = "NO DATA";

                    }

                    VoiceAttackPlugin.LogOutput("OriginTemp", "grey");

                }
                else
                {

                    OriginTemp = "NO DATA";

                }


            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("OriginTemp. " + e.Message, "red");
            }

            try
            {
                Altn = sbOFP["alternate"]["icao_code"].InnerText;
                VoiceAttackPlugin.LogOutput("Altn", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("Altn. " + e.Message, "red");
            }

            try
            {
                AltnElevation = Convert.ToInt32(sbOFP["alternate"]["elevation"].InnerText);
                VoiceAttackPlugin.LogOutput("AltnElevation", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("AltnElevation. " + e.Message, "red");
            }

            try
            {
                AltnRwy = sbOFP["alternate"]["plan_rwy"].InnerText;
                VoiceAttackPlugin.LogOutput("AltnRwy", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("AltnRwy. " + e.Message, "red");
            }

            try
            {
                AltnTransAlt = Convert.ToInt32(sbOFP["alternate"]["trans_alt"].InnerText);
                VoiceAttackPlugin.LogOutput("AltnTransAlt", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("AltnTransAlt. " + e.Message, "red");
            }

            try
            {
                AltnTransLevel = Convert.ToInt32(sbOFP["alternate"]["trans_level"].InnerText);
                VoiceAttackPlugin.LogOutput("AltnTransLevel", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("AltnTransLevel. " + e.Message, "red");
            }

            try
            {
                AltnMetar = sbOFP["alternate"]["metar"].InnerText;
                VoiceAttackPlugin.LogOutput("AltnMetar", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("AltnMetar. " + e.Message, "red");
            }

            try
            {
                AltnMetarTime = sbOFP["alternate"]["metar_time"].InnerText;
                VoiceAttackPlugin.LogOutput("AltnMetarTime", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("AltnMetarTime. " + e.Message, "red");
            }

            try
            {
                AltnTAF = sbOFP["alternate"]["taf"].InnerText;
                VoiceAttackPlugin.LogOutput("AltnTAF", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("AltnTAF. " + e.Message, "red");
            }

            try
            {
                AltnTAFTime = sbOFP["alternate"]["taf_time"].InnerText;
                VoiceAttackPlugin.LogOutput("AltnTAFTime", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("AltnTAFTime. " + e.Message, "red");
            }


            try
            {
                if (AltnTAFTime.Length > 1 && AltnTAF.IndexOf("KT ") != -1)
                {
                    Utils.sbAltnAvailWeather = 2;
                    Utils.sbWeatherReportCharPos = AltnTAF.IndexOf("KT ");
                    while (AltnTAF.Substring(Utils.sbWeatherReportCharPos, 1) != " ")
                    {
                        Utils.sbWeatherReportCharPos--;
                    }
                    Utils.sbWeatherReportCharPos++;

                    AltnWind = AltnTAF.Substring(Utils.sbWeatherReportCharPos, AltnTAF.IndexOf("KT ") + 2 - Utils.sbWeatherReportCharPos);
                    VoiceAttackPlugin.LogOutput("AltnWind", "grey");

                }
                else if (AltnMetarTime.Length > 1 && AltnMetar.IndexOf("KT ") != -1)
                {
                    Utils.sbAltnAvailWeather = 1;
                    Utils.sbWeatherReportCharPos = AltnMetar.IndexOf("KT ");
                    while (AltnMetar.Substring(Utils.sbWeatherReportCharPos, 1) != " ")
                    {
                        Utils.sbWeatherReportCharPos--;
                    }
                    Utils.sbWeatherReportCharPos++;

                    AltnWind = AltnMetar.Substring(Utils.sbWeatherReportCharPos, AltnMetar.IndexOf("KT ") + 2 - Utils.sbWeatherReportCharPos);
                    VoiceAttackPlugin.LogOutput("AltnWind", "grey");

                }
                else
                {
                    Utils.sbAltnAvailWeather = 0;
                    AltnWind = "NO DATA";
                }
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("AltnWind. " + e.Message, "red");
            }

            try
            {

                double presHPA = 0;
                double presINHG = 0;
                string presHPAs = "NULL";
                string presINHGs = "NULL";

                if (AltnMetarTime.Length > 1)
                {

                    if (AltnMetar.IndexOf(" Q") != -1)
                    {

                        Utils.sbWeatherReportCharPos = AltnMetar.IndexOf(" Q");
                        AltnPressure = AltnMetar.Substring(Utils.sbWeatherReportCharPos + 1, 5);

                    }

                    else if (AltnMetar.IndexOf(" A") != -1)
                    {

                        Utils.sbWeatherReportCharPos = AltnMetar.IndexOf(" A");
                        AltnPressure = AltnMetar.Substring(Utils.sbWeatherReportCharPos + 1, 5);

                    }

                    else
                    {

                        AltnPressure = "NO DATA";

                    }

                    if (AltnPressure.Substring(0,1) == "Q")
                    {
                        presHPAs = AltnPressure.Substring(1,4);
                        presHPA = Convert.ToDouble(presHPAs);

                        presINHG = presHPA / 33.864;
                        AltnPressureINHG = presINHG;
                        presINHGs = Convert.ToString(presINHG);
                        presINHGs = presINHGs.Substring(0, 5);

                        AltnPressure = presHPAs + " hPa / " + presINHGs + " inHg";

                    }
                    else if (AltnPressure.Substring(0,1) == "A")
                    {
                        presINHGs = AltnPressure.Substring(1, 4);
                        presINHG = Convert.ToDouble(presINHGs);
                        presINHG = presINHG / 100;
                        AltnPressureINHG = presINHG;

                        presHPA = presINHG * 33.864;
                        presHPAs = Convert.ToString(presHPA);
                        presHPAs = presHPAs.Substring(0, 4);

                        AltnPressure = presHPAs + " hPa / " + presINHGs + " inHg";
                    }

                    VoiceAttackPlugin.LogOutput("AltnPressure", "grey");

                }
                else
                {

                    AltnPressure = "NO DATA";

                }


            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("AltnPressure. " + e.Message, "red");
            }

            try
            {
                if (AltnMetarTime.Length > 1)
                {

                    if (AltnMetar.IndexOf("/") != -1 && (AltnMetar.Substring(AltnMetar.IndexOf("/") + 3, 1) == " " || AltnMetar.Substring(AltnMetar.IndexOf("/") + 4, 1) == " "))
                    {

                        Utils.sbWeatherReportCharPos = AltnMetar.IndexOf("/");
                        while (AltnMetar.Substring(Utils.sbWeatherReportCharPos,1) != " ")
                        {
                            Utils.sbWeatherReportCharPos--;
                        }
                        Utils.sbWeatherReportCharPos++;
                        Utils.sbWeatherReportCharPos2 = Utils.sbWeatherReportCharPos;
                        while (AltnMetar.Substring(Utils.sbWeatherReportCharPos,1) != " ")
                        {
                            Utils.sbWeatherReportCharPos++;
                        }
                        Utils.sbWeatherReportCharPos--;
                        AltnTemp = AltnMetar.Substring(Utils.sbWeatherReportCharPos2, Utils.sbWeatherReportCharPos+1 - Utils.sbWeatherReportCharPos2);

                    }

                    else
                    {

                        AltnTemp = "NO DATA";

                    }

                    VoiceAttackPlugin.LogOutput("AltnTemp", "grey");

                }
                else
                {

                    OriginTemp = "NO DATA";

                }


            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("AltnTemp. " + e.Message, "red");
            }

            try
            {
                Destination = sbOFP["destination"]["icao_code"].InnerText;
                VoiceAttackPlugin.LogOutput("Destination", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("Destination. " + e.Message, "red");
            }

            try
            {
                DestElevation = Convert.ToInt32(sbOFP["destination"]["elevation"].InnerText);
                VoiceAttackPlugin.LogOutput("DestElevation", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("DestElevation. " + e.Message, "red");
            }

            try
            {
                DestRwy = sbOFP["destination"]["plan_rwy"].InnerText;
                VoiceAttackPlugin.LogOutput("DestRwy", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("DestRwy. " + e.Message, "red");
            }

            try
            {
                DestTransAlt = Convert.ToInt32(sbOFP["destination"]["trans_alt"].InnerText);
                VoiceAttackPlugin.LogOutput("DestTransAlt", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("DestTransAlt. " + e.Message, "red");
            }

            try
            {
                DestTransLevel = Convert.ToInt32(sbOFP["destination"]["trans_level"].InnerText);
                VoiceAttackPlugin.LogOutput("DestTransLevel", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("DestTransLevel. " + e.Message, "red");
            }

            try
            {
                DestMetar = sbOFP["destination"]["metar"].InnerText;
                VoiceAttackPlugin.LogOutput("t", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("DestMetar. " + e.Message, "red");
            }

            try
            {
                DestMetarTime = sbOFP["destination"]["metar_time"].InnerText;
                VoiceAttackPlugin.LogOutput("DestMetarTime", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("DestMetarTime. " + e.Message, "red");
            }

            try
            {
                DestTAF = sbOFP["destination"]["taf"].InnerText;
                VoiceAttackPlugin.LogOutput("DestTAF", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("DestTAF. " + e.Message, "red");
            }

            try
            {
                DestTAFTime = sbOFP["destination"]["taf_time"].InnerText;
                VoiceAttackPlugin.LogOutput("DestTAFTime", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("DestMetarTime. " + e.Message, "red");
            }


            try
            {
                if (DestTAFTime.Length > 1 && DestTAF.IndexOf("KT ") != -1)
                {
                    Utils.sbDestAvailWeather = 2;
                    Utils.sbWeatherReportCharPos = DestTAF.IndexOf("KT ");
                    while (DestTAF.Substring(Utils.sbWeatherReportCharPos, 1) != " ")
                    {
                        Utils.sbWeatherReportCharPos--;
                    }
                    Utils.sbWeatherReportCharPos++;

                    DestWind = DestTAF.Substring(Utils.sbWeatherReportCharPos, DestTAF.IndexOf("KT ") + 2 - Utils.sbWeatherReportCharPos);
                    VoiceAttackPlugin.LogOutput("DestWind", "grey");

                }
                else if (DestMetarTime.Length > 1 && DestMetar.IndexOf("KT ") != -1)
                {
                    Utils.sbDestAvailWeather = 1;
                    Utils.sbWeatherReportCharPos = DestMetar.IndexOf("KT ");
                    while (DestMetar.Substring(Utils.sbWeatherReportCharPos, 1) != " ")
                    {
                        Utils.sbWeatherReportCharPos--;
                    }
                    Utils.sbWeatherReportCharPos++;

                    DestWind = DestMetar.Substring(Utils.sbWeatherReportCharPos, DestMetar.IndexOf("KT ") + 2 - Utils.sbWeatherReportCharPos);
                    VoiceAttackPlugin.LogOutput("DestWind", "grey");

                }
                else
                {
                    Utils.sbDestAvailWeather = 0;
                    DestWind = "NO DATA";
                }
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("DestWind. " + e.Message, "red");
            }


            try
            {
                double presHPA = 0;
                double presINHG = 0;
                string presHPAs = "NULL";
                string presINHGs = "NULL";

                if (DestMetarTime.Length > 1)
                {

                    if (DestMetar.IndexOf(" Q") != -1)
                    {

                        Utils.sbWeatherReportCharPos = DestMetar.IndexOf(" Q");
                        DestPressure = DestMetar.Substring(Utils.sbWeatherReportCharPos + 1, 5);

                    }

                    else if (DestMetar.IndexOf(" A") != -1)
                    {

                        Utils.sbWeatherReportCharPos = DestMetar.IndexOf(" A");
                        DestPressure = DestMetar.Substring(Utils.sbWeatherReportCharPos + 1, 5);

                    }

                    else
                    {

                        DestPressure = "NO DATA";

                    }

                    if (DestPressure.Substring(0, 1) == "Q")
                    {
                        presHPAs = DestPressure.Substring(1, 4);
                        presHPA = Convert.ToDouble(presHPAs);

                        presINHG = presHPA / 33.864;
                        DestPressureINHG = presINHG;
                        presINHGs = Convert.ToString(presINHG);
                        presINHGs = presINHGs.Substring(0, 5);

                        DestPressure = presHPAs + " hPa / " + presINHGs + " inHg";

                    }
                    else if (AltnPressure.Substring(0, 1) == "A")
                    {
                        presINHGs = DestPressure.Substring(1, 4);
                        presINHG = Convert.ToDouble(presINHGs);
                        presINHG = presINHG / 100;
                        DestPressureINHG = presINHG;

                        presHPA = presINHG * 33.864;
                        presHPAs = Convert.ToString(presHPA);
                        presHPAs = presHPAs.Substring(0, 4);

                        DestPressure = presHPAs + " hPa / " + presINHGs + " inHg";
                    }

                    VoiceAttackPlugin.LogOutput("DestPressure", "grey");

                }

                else
                {

                    DestPressure = "NO DATA";

                }
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("DestPressure. " + e.Message, "red");
            }

            try
            {
                if (DestMetarTime.Length > 1)
                {

                    if (DestMetar.IndexOf("/") != -1 && (DestMetar.Substring(DestMetar.IndexOf("/") + 3, 1) == " " || DestMetar.Substring(DestMetar.IndexOf("/") + 4, 1) == " "))
                    {

                        Utils.sbWeatherReportCharPos = DestMetar.IndexOf("/");
                        while (DestMetar.Substring(Utils.sbWeatherReportCharPos, 1) != " ")
                        {
                            Utils.sbWeatherReportCharPos--;
                        }
                        Utils.sbWeatherReportCharPos++;
                        Utils.sbWeatherReportCharPos2 = Utils.sbWeatherReportCharPos;
                        while (DestMetar.Substring(Utils.sbWeatherReportCharPos, 1) != " ")
                        {
                            Utils.sbWeatherReportCharPos++;
                        }
                        Utils.sbWeatherReportCharPos--;
                        DestTemp = DestMetar.Substring(Utils.sbWeatherReportCharPos2, Utils.sbWeatherReportCharPos+1 - Utils.sbWeatherReportCharPos2);

                    }

                    else
                    {

                        DestTemp = "NO DATA";

                    }

                    VoiceAttackPlugin.LogOutput("DestTemp", "grey");

                }
                else
                {

                    DestTemp = "NO DATA";

                }


            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("DestTemp. " + e.Message, "red");
            }

            try
            {
                Units = sbOFP["params"]["units"].InnerText;
                VoiceAttackPlugin.LogOutput("Units", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("Units. " + e.Message, "red");
            }

            try
            {
                FinRes = Convert.ToDouble(sbOFP["fuel"]["reserve"].InnerText);
                VoiceAttackPlugin.LogOutput("FinRes", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("FinRes. " + e.Message, "red");
            }

            try
            {
                AltnFuel = Convert.ToDouble(sbOFP["fuel"]["alternate_burn"].InnerText);
                VoiceAttackPlugin.LogOutput("AltnFuel", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("AltnFuel. " + e.Message, "red");
            }

            try
            {
                FinresPAltn = Convert.ToDouble(sbOFP["fuel"]["reserve"].InnerText) + Convert.ToDouble(sbOFP["fuel"]["alternate_burn"].InnerText);
                VoiceAttackPlugin.LogOutput("FinresPAltn", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("FinresPAltn. " + e.Message, "red");
            }

            try
            {
                Fuel = Convert.ToDouble(sbOFP["fuel"]["plan_ramp"].InnerText);
                VoiceAttackPlugin.LogOutput("Fuel", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("Fuel. " + e.Message, "red");
            }

            try
            {
                Passenger = Convert.ToInt32(sbOFP["weights"]["pax_count"].InnerText);
                VoiceAttackPlugin.LogOutput("Passenger", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("Passenger. " + e.Message, "red");
            }

            try
            {
                Bags = Convert.ToInt32(sbOFP["weights"]["bag_count"].InnerText);
                VoiceAttackPlugin.LogOutput("Bags", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("Bags. " + e.Message, "red");
            }

            try
            {
                WeightPax = Convert.ToDouble(sbOFP["weights"]["payload"].InnerText) - Convert.ToDouble(sbOFP["weights"]["cargo"].InnerText);
                VoiceAttackPlugin.LogOutput("WeightPax", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("WeightPax. " + e.Message, "red");
            }

            try
            {
                WeightCargo = Convert.ToDouble(sbOFP["weights"]["cargo"].InnerText);
                VoiceAttackPlugin.LogOutput("WeightCargo", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("WeightCargo. " + e.Message, "red");
            }

            try
            {
                Payload = Convert.ToDouble(sbOFP["weights"]["payload"].InnerText);
                VoiceAttackPlugin.LogOutput("Payload", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("Payload. " + e.Message, "red");
            }

            try
            {
                ZFW = Convert.ToDouble(sbOFP["weights"]["est_zfw"].InnerText);
                VoiceAttackPlugin.LogOutput("ZFW", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("ZFW. " + e.Message, "red");
            }

            try
            {
                TOW = Convert.ToDouble(sbOFP["weights"]["est_tow"].InnerText);
                VoiceAttackPlugin.LogOutput("TOW", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogMonitorOutput("TOW. " + e.Message, "red");
            }



            Utils.sbFlight = Flight;
            Utils.sbAirlineICAO = AirlineICAO;
            Utils.sbCruiseProf = CruiseProf;
            Utils.sbClimbProf = ClimbProf;
            Utils.sbDescentProf = DescentProf;
            Utils.sbCostIndex = CostIndex;
            Utils.sbInitialAlt = InitialAlt;
            Utils.sbAvgWindComp = AvgWindComp;
            Utils.sbAvgWindDir = AvgWindDir;
            Utils.sbAvgWindSpd = AvgWindSpd;
            Utils.sbTopClimbOAT = TopClimbOAT;
            Utils.sbRoute = Route;
            Utils.sbOrigin = Origin;
            Utils.sbOriginElevation = OriginElevation;
            Utils.sbOriginRwy = OriginRwy;
            Utils.sbOriginTransAlt = OriginTransAlt;
            Utils.sbOriginTransLevel = OriginTransLevel;
            Utils.sbOriginWind = OriginWind;
            Utils.sbOriginMetar = OriginMetar;
            Utils.sbOriginTAF = OriginTAF;
            Utils.sbOriginPressure = OriginPressure;
            Utils.sbOriginPressureINHG= OriginPressureINHG;
            Utils.sbOriginTemp = OriginTemp;
            Utils.sbAltn = Altn;
            Utils.sbAltnElevation = AltnElevation;
            Utils.sbAltnRwy = AltnRwy;
            Utils.sbAltnTransAlt = AltnTransAlt;
            Utils.sbAltnTransLevel = AltnTransLevel;
            Utils.sbAltnWind = AltnWind;
            Utils.sbAltnMetar = AltnMetar;
            Utils.sbAltnTAF = AltnTAF;
            Utils.sbAltnPressure = AltnPressure;
            Utils.sbAltnPressureINHG = AltnPressureINHG;
            Utils.sbAltnTemp = AltnTemp;
            Utils.sbDestination = Destination;
            Utils.sbDestElevation = DestElevation;
            Utils.sbDestRwy = DestRwy;
            Utils.sbDestTransAlt = DestTransAlt;
            Utils.sbDestTransLevel = DestTransLevel;
            Utils.sbDestWind = DestWind;
            Utils.sbDestMetar = DestMetar;
            Utils.sbDestTAF = DestTAF;
            Utils.sbDestPressure = DestPressure;
            Utils.sbDestPressureINHG = DestPressureINHG;
            Utils.sbDestTemp = DestTemp;
            Utils.sbUnits = Units;
            Utils.sbFinRes = FinRes;
            Utils.sbAltnFuel = AltnFuel;
            Utils.sbFinresPAltn = FinresPAltn;
            Utils.sbFuel = Fuel;
            Utils.sbPassenger = Passenger;
            Utils.sbBags = Bags;
            Utils.sbWeightPax = WeightPax;
            Utils.sbWeightCargo = WeightCargo;
            Utils.sbPayload = Payload;
            Utils.sbZFW = ZFW;
            Utils.sbTOW = TOW;
            


            VoiceAttackPlugin.LogOutput("Receiving OFP for Flight " + Flight + ". (" + Origin + " -> " + Destination + ")", "grey");


            return true;
        }




            


     
    }
}
