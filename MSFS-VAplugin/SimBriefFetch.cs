using System;
using System.Collections.Generic;
using System.Linq;
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

namespace MSFS
{
    public class SimBriefFetch
    {

        public string Flight { get; set; } = "";
        public int CostIndex { get; set; }
        public string Route { get; set; }
        public string Origin { get; set; }
        public int OriginElevation { get; set; }
        public string OriginRwy { get; set; }
        public int OriginTransAlt { get; set; }
        public int OriginTransLevel { get; set; }
        public string OriginMetar { get; set; }
        public int OriginWindDir { get; set; }
        public int OriginWindSpd { get; set; }
        public int OriginQNH { get; set; }
        public double OriginBaro { get; set; }
        public string Destination { get; set; }
        public int DestElevation { get; set; }
        public string DestRwy { get; set; }
        public int DestTransAlt { get; set; }
        public int DestTransLevel { get; set; }
        public string DestMetar { get; set; }
        public int DestWindDir { get; set; }
        public int DestWindSpd { get; set; }
        public int DestQNH { get; set; }
        public double DestBaro { get; set; }
        public string Altn { get; set; }
        public int AltnElevation { get; set; }
        public string AltnRwy { get; set; }
        public int AltnTransAlt { get; set; }
        public int AltnTransLevel { get; set; }
        public string AltnMetar { get; set; }
        public int AltnWindDir { get; set; }
        public int AltnWindSpd { get; set; }
        public int AltnQNH { get; set; }
        public double AltnBaro { get; set; }
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

        protected XmlNode LoadFile()
        {


            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(File.ReadAllText(Utils.xmlFile));
            VoiceAttackPlugin.LogOutput("XML File parsed");
            return xmlDoc.ChildNodes[0];
        }

        protected XmlNode LoadOFP()
        {

            
            string a = "Loading OFP..." + Utils.simbriefURL + Utils.pilotID;

            VoiceAttackPlugin.LogOutput(a, "grey");
            if (!File.Exists(Utils.xmlFile))
            {
                VoiceAttackPlugin.LogOutput("OFP file does not exist. Fetching.", "grey");
                return FetchOnline();
            }

            else
            {
                VoiceAttackPlugin.LogOutput("OFP file exists. Loading.", "grey");
                return LoadFile();
            }
                
        }

        public bool Load()
        {
            XmlNode sbOFP = LoadOFP();
            VoiceAttackPlugin.LogOutput("Importing data...", "grey");

            Flight = sbOFP["general"]["icao_airline"].InnerText + sbOFP["general"]["flight_number"].InnerText;
            VoiceAttackPlugin.LogOutput("Flight", "grey");

            CostIndex = Convert.ToInt32(sbOFP["general"]["costindex"].InnerText);
            VoiceAttackPlugin.LogOutput("CostIndex", "grey");

            Route = sbOFP["general"]["route"].InnerText;
            VoiceAttackPlugin.LogOutput("Route", "grey");

            Origin = sbOFP["origin"]["icao_code"].InnerText;
            VoiceAttackPlugin.LogOutput("Origin", "grey");

            OriginElevation = Convert.ToInt32(sbOFP["origin"]["elevation"].InnerText);
            VoiceAttackPlugin.LogOutput("OriginElevation", "grey");

            OriginRwy = sbOFP["origin"]["plan_rwy"].InnerText;
            VoiceAttackPlugin.LogOutput("OriginRwy", "grey");

            OriginTransAlt = Convert.ToInt32(sbOFP["origin"]["trans_alt"].InnerText);
            VoiceAttackPlugin.LogOutput("OriginTransAlt", "grey");

            OriginTransLevel = Convert.ToInt32(sbOFP["origin"]["trans_level"].InnerText);
            VoiceAttackPlugin.LogOutput("OriginTransLevel", "grey");

            OriginMetar = sbOFP["origin"]["metar"].InnerText;
            VoiceAttackPlugin.LogOutput("OriginMetar", "grey");

            OriginWindDir = Convert.ToInt32(OriginMetar.Substring(OriginMetar.IndexOf("KT ") - 5, 3));
            VoiceAttackPlugin.LogOutput("OriginWindDir", "grey");

            OriginWindSpd = Convert.ToInt32(OriginMetar.Substring(OriginMetar.IndexOf("KT ") - 2, 2));
            VoiceAttackPlugin.LogOutput("OriginWindSpd", "grey");

            OriginQNH = Convert.ToInt32(OriginMetar.Substring(OriginMetar.IndexOf(" Q") + 2, 4));
            VoiceAttackPlugin.LogOutput("OriginQNH", "grey");

            OriginBaro = OriginQNH/33.864;
            VoiceAttackPlugin.LogOutput("OriginBaro", "grey");

            Altn = sbOFP["alternate"]["icao_code"].InnerText;
            VoiceAttackPlugin.LogOutput("Altn", "grey");

            AltnElevation = Convert.ToInt32(sbOFP["alternate"]["elevation"].InnerText);
            VoiceAttackPlugin.LogOutput("AltnElevation", "grey");

            AltnRwy = sbOFP["alternate"]["plan_rwy"].InnerText;
            VoiceAttackPlugin.LogOutput("AltnRwy", "grey");

            AltnTransAlt = Convert.ToInt32(sbOFP["alternate"]["trans_alt"].InnerText);
            VoiceAttackPlugin.LogOutput("AltnTransAlt", "grey");

            AltnTransLevel = Convert.ToInt32(sbOFP["alternate"]["trans_level"].InnerText);
            VoiceAttackPlugin.LogOutput("AltnTransLevel", "grey");

            AltnMetar = sbOFP["alternate"]["metar"].InnerText;
            VoiceAttackPlugin.LogOutput("AltnMetar", "grey");

            AltnWindDir = Convert.ToInt32(AltnMetar.Substring(AltnMetar.IndexOf("KT ") - 5, 3));
            VoiceAttackPlugin.LogOutput("AltnWindDir", "grey");

            AltnWindSpd = Convert.ToInt32(AltnMetar.Substring(AltnMetar.IndexOf("KT ") - 2, 2));
            VoiceAttackPlugin.LogOutput("AltnWindSpd", "grey");

            AltnQNH = Convert.ToInt32(AltnMetar.Substring(AltnMetar.IndexOf(" Q") + 2, 4));
            VoiceAttackPlugin.LogOutput("AltnQNH", "grey");

            AltnBaro = AltnQNH / 33.864;
            VoiceAttackPlugin.LogOutput("AltnBaro", "grey");

            Destination = sbOFP["destination"]["icao_code"].InnerText;
            VoiceAttackPlugin.LogOutput("Destination", "grey");

            DestElevation = Convert.ToInt32(sbOFP["destination"]["elevation"].InnerText);
            VoiceAttackPlugin.LogOutput("DestElevation", "grey");

            DestRwy = sbOFP["destination"]["plan_rwy"].InnerText;
            VoiceAttackPlugin.LogOutput("DestRwy", "grey");

            DestTransAlt = Convert.ToInt32(sbOFP["destination"]["trans_alt"].InnerText);
            VoiceAttackPlugin.LogOutput("DestTransAlt", "grey");

            DestTransLevel = Convert.ToInt32(sbOFP["destination"]["trans_level"].InnerText);
            VoiceAttackPlugin.LogOutput("DestTransLevel", "grey");

            DestMetar = sbOFP["destination"]["metar"].InnerText;
            VoiceAttackPlugin.LogOutput("t", "grey");

            DestWindDir = Convert.ToInt32(DestMetar.Substring(DestMetar.IndexOf("KT ") - 5, 3));
            VoiceAttackPlugin.LogOutput("DestWindDir", "grey");

            DestWindSpd = Convert.ToInt32(DestMetar.Substring(DestMetar.IndexOf("KT ") - 2, 2));
            VoiceAttackPlugin.LogOutput("DestWindSpd", "grey");

            DestQNH = Convert.ToInt32(DestMetar.Substring(DestMetar.IndexOf(" Q") + 2, 4));
            VoiceAttackPlugin.LogOutput("DestQNH", "grey");

            DestBaro = DestQNH / 33.864;
            VoiceAttackPlugin.LogOutput("DestBaro", "grey");

            Units = sbOFP["params"]["units"].InnerText;
            VoiceAttackPlugin.LogOutput("Units", "grey");

            FinRes = Convert.ToDouble(sbOFP["fuel"]["reserve"].InnerText);
            VoiceAttackPlugin.LogOutput("FinRes", "grey");

            AltnFuel = Convert.ToDouble(sbOFP["fuel"]["alternate_burn"].InnerText);
            VoiceAttackPlugin.LogOutput("AltnFuel", "grey");

            FinresPAltn = Convert.ToDouble(sbOFP["fuel"]["reserve"].InnerText) + Convert.ToDouble(sbOFP["fuel"]["alternate_burn"].InnerText);
            VoiceAttackPlugin.LogOutput("FinresPAltn", "grey");

            Fuel = Convert.ToDouble(sbOFP["fuel"]["plan_ramp"].InnerText);
            VoiceAttackPlugin.LogOutput("Fuel", "grey");

            Passenger = Convert.ToInt32(sbOFP["weights"]["pax_count"].InnerText);
            VoiceAttackPlugin.LogOutput("Passenger", "grey");

            Bags = Convert.ToInt32(sbOFP["weights"]["bag_count"].InnerText);
            VoiceAttackPlugin.LogOutput("Bags", "grey");

            WeightPax = Convert.ToDouble(sbOFP["weights"]["payload"].InnerText) - Convert.ToDouble(sbOFP["weights"]["cargo"].InnerText);
            VoiceAttackPlugin.LogOutput("WeightPax", "grey");

            WeightCargo = Convert.ToDouble(sbOFP["weights"]["cargo"].InnerText);
            VoiceAttackPlugin.LogOutput("WeightCargo", "grey");

            Payload = Convert.ToDouble(sbOFP["weights"]["payload"].InnerText);
            VoiceAttackPlugin.LogOutput("Payload", "grey");

            ZFW = Convert.ToDouble(sbOFP["weights"]["est_zfw"].InnerText);
            VoiceAttackPlugin.LogOutput("ZFW", "grey");

            TOW = Convert.ToDouble(sbOFP["weights"]["est_tow"].InnerText);
            VoiceAttackPlugin.LogOutput("TOW", "grey");


            Utils.sbFlight = Flight;
            Utils.sbCostIndex = CostIndex;
            Utils.sbRoute = Route;
            Utils.sbOrigin = Origin;
            Utils.sbOriginElevation = OriginElevation;
            Utils.sbOriginRwy = OriginRwy;
            Utils.sbOriginTransAlt = OriginTransAlt;
            Utils.sbOriginTransLevel = OriginTransLevel;
            Utils.sbOriginWindDir = OriginWindDir;
            Utils.sbOriginWindSpd = OriginWindSpd;
            Utils.sbOriginMetar = OriginMetar;
            Utils.sbOriginQNH = OriginQNH;
            Utils.sbOriginBaro = OriginBaro;
            Utils.sbAltn = Altn;
            Utils.sbAltnElevation = AltnElevation;
            Utils.sbAltnRwy = AltnRwy;
            Utils.sbAltnTransAlt = AltnTransAlt;
            Utils.sbAltnTransLevel = AltnTransLevel;
            Utils.sbAltnWindDir = AltnWindDir;
            Utils.sbAltnWindSpd = AltnWindSpd;
            Utils.sbAltnMetar = AltnMetar;
            Utils.sbAltnQNH = AltnQNH;
            Utils.sbAltnBaro = AltnBaro;
            Utils.sbDestination = Destination;
            Utils.sbDestElevation = DestElevation;
            Utils.sbDestRwy = DestRwy;
            Utils.sbDestTransAlt = DestTransAlt;
            Utils.sbDestTransLevel = DestTransLevel;
            Utils.sbDestWindDir = DestWindDir;
            Utils.sbDestWindSpd = DestWindSpd;
            Utils.sbDestMetar = DestMetar;
            Utils.sbDestQNH = DestQNH;
            Utils.sbDestBaro = DestBaro;
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
