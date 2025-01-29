using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Composition.Primitives;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using FSUIPC;

namespace MSFS
{
    public class WeatherReport
    {
        bool stationIdentified;
        bool wpExists;
        bool scExists;
        bool rvrExists;
        bool stExists;
        bool intVisExists;
        bool windVariableRange;
        bool windHasGust;
        bool voiceAttackVariablesSave = false;

        string reportPrefix = "";
        string reportSufix = "";
        string stationIdentifier = "";
        string observationTime = "";
        string windDirection = "";  
        string windSpeed = "";
        string windGust = "";
        string windUnits = "";
        string windType = "";
        string windFrom = "";
        string windTo = "";
        string visibilityValueKM = "";
        string visibilityValueSM = "";
        string stRunway = "";
        string stDeposit = "";
        string stContamination = "";
        string stDepth = "";
        string stBraking = "";
        string rvrTrend = "";
        string rvrVisibility = "";
        string rvrRunway = "";
        string rvrUnit = "";
        string finalWeatherMessage = "";
        string tfcstMinMessage = "";
        string tfcstMaxMessage = "";
        string wsalt = "";
        string wsdir = "";
        string wsspd = "";
        string wsunit = "";
        string wsmessage = "";
        int intVisValue = 0;
        int minTemperatureCount = 0;
        int maxTemperatureCount = 0;

        List<string> wpParts = new List<string>();
        List<string> scParts = new List<string>();
        List<string> rvrParts = new List<string>();
        List<string> stParts = new List<string>();
        List<string> tafSections = new List<string>();
        List<string> interSections = new List<string>();
        List<string> afterInterSections = new List<string>();

        int weatherPhenomenaIndex = 0;
        int skyConditionsIndex = 0;
        int remarksIndex = 0;
        int rvrIndex = -1;
        int snowtamIndex = -1;
        int sectionIndex = 0;

        //METAR PARSING
        public void MetarReport(string metarRep)
        {
            windVariableRange = false;
            windHasGust = false;
            stationIdentified = false;
            wpExists = false;
            scExists = false;
            rvrExists = false;
            intVisExists = false;
            Utils.tempoRepExists = false;
            Utils.becmgRepExists = false;
            Utils.rmkRepExists = false;
            VoiceAttackPlugin.SetText("RMK.Exists", "false");
            VoiceAttackPlugin.SetText("TEMPO.Exists", "false");
            VoiceAttackPlugin.SetText("BECMG.Exists", "false");

            Utils.wpQuantity = 0;

            stationIdentifier = "";
            observationTime = "";
            windDirection = "";
            windSpeed = "";
            windGust = "";
            windUnits = "";
            windType = "";
            windFrom = "";
            windTo = "";
            visibilityValueKM = "";
            visibilityValueSM = "";
            stRunway = "";
            stDeposit = "";
            stContamination = "";
            stDepth = "";
            stBraking = "";
            rvrTrend = "";
            rvrVisibility = "";
            rvrRunway = "";
            rvrUnit = "";
            finalWeatherMessage = "";
            wsalt = "";
            wsdir = "";
            wsspd = "";
            wsunit = "";
            wsmessage = "";
            intVisValue = 0;

            wpParts.Clear();
            scParts.Clear();
            rvrParts.Clear();
            stParts.Clear();

            reportPrefix = "METAR";
            reportSufix = "";

            rvrIndex = -1;
            snowtamIndex = -1;
            voiceAttackVariablesSave = true;

            int becmgIndex = metarRep.IndexOf("BECMG");
            int tempoIndex = metarRep.IndexOf("TEMPO");
            int rmkIndex = metarRep.IndexOf("RMK");

            metarRep = Regex.Replace(metarRep, @"\b(METAR|SPECI|AUTO|TAF)\b\s*", "");
            VoiceAttackPlugin.LogOutput("Removed prefix from METAR: " + metarRep, "grey");

            if (rmkIndex >= 0)
            {
                Utils.rmkRep = metarRep.Substring(rmkIndex).Trim();
                metarRep = metarRep.Substring(0, rmkIndex).Trim();
                Utils.rmkRepExists = true;
                VoiceAttackPlugin.SetText("RMK.Exists", "true");
                VoiceAttackPlugin.LogOutput("RMK section in METAR report detected: " + Utils.rmkRep, "grey");
                VoiceAttackPlugin.LogOutput("Trimmed METAR report: " + metarRep, "grey");
            }

            if (tempoIndex >= 0)
            {
                if (becmgIndex >= 0)
                {
                    if (becmgIndex > tempoIndex)
                    {
                        VoiceAttackPlugin.LogOutput("TEMPO and BECMG sections detected", "grey");

                        Utils.becmgRep = metarRep.Substring(becmgIndex).Trim();
                        Utils.becmgRep = Regex.Replace(Utils.becmgRep, @"\b(BECMG)\b\s*", "");
                        Utils.becmgRep = Utils.becmgRep.Trim();
                        Utils.becmgRep = "BECMG " + Utils.becmgRep;
                        metarRep = metarRep.Substring(0, becmgIndex).Trim();
                        metarRep = Regex.Replace(metarRep, @"\b(BECMG)\b\s*", "");
                        metarRep = metarRep.Trim();
                        Utils.becmgRepExists = true;

                        Utils.tempoRep = metarRep.Substring(tempoIndex).Trim();
                        Utils.tempoRep = Regex.Replace(Utils.tempoRep, @"\b(TEMPO)\b\s*", "");
                        Utils.tempoRep = Utils.tempoRep.Trim();
                        Utils.tempoRep = "TEMPO " + Utils.tempoRep;
                        metarRep = metarRep.Substring(0, tempoIndex).Trim();
                        metarRep = Regex.Replace(metarRep, @"\b(TEMPO)\b\s*", "");
                        metarRep = metarRep.Trim();
                        Utils.tempoRepExists = true;

                        VoiceAttackPlugin.SetText("TEMPO.Exists", "true");
                        VoiceAttackPlugin.SetText("BECMG.Exists", "true");
                        VoiceAttackPlugin.LogOutput("TEMPO section in METAR report detected: " + Utils.tempoRep, "grey");
                        VoiceAttackPlugin.LogOutput("BECMG section in METAR report detected: " + Utils.becmgRep, "grey");
                        VoiceAttackPlugin.LogOutput("Trimmed METAR report: " + metarRep, "grey");
                    }
                    else
                    {
                        VoiceAttackPlugin.LogOutput("BECMG and TEMPO sections detected", "grey");

                        Utils.tempoRep = metarRep.Substring(tempoIndex).Trim();
                        Utils.tempoRep = Regex.Replace(Utils.tempoRep, @"\b(TEMPO)\b\s*", "");
                        Utils.tempoRep = Utils.tempoRep.Trim();
                        Utils.tempoRep = "TEMPO " + Utils.tempoRep;
                        metarRep = metarRep.Substring(0, tempoIndex).Trim();
                        metarRep = Regex.Replace(metarRep, @"\b(TEMPO)\b\s*", "");
                        metarRep = metarRep.Trim();
                        Utils.tempoRepExists = true;

                        Utils.becmgRep = metarRep.Substring(becmgIndex).Trim();
                        Utils.becmgRep = Regex.Replace(Utils.becmgRep, @"\b(BECMG)\b\s*", "");
                        Utils.becmgRep = Utils.becmgRep.Trim();
                        Utils.becmgRep = "BECMG " + Utils.becmgRep;
                        metarRep = metarRep.Substring(0, becmgIndex).Trim();
                        metarRep = Regex.Replace(metarRep, @"\b(BECMG)\b\s*", "");
                        metarRep = metarRep.Trim();
                        Utils.becmgRepExists = true;

                        VoiceAttackPlugin.SetText("TEMPO.Exists", "true");
                        VoiceAttackPlugin.SetText("BECMG.Exists", "true");
                        VoiceAttackPlugin.LogOutput("TEMPO section in METAR report detected: " + Utils.tempoRep, "grey");
                        VoiceAttackPlugin.LogOutput("BECMG section in METAR report detected: " + Utils.becmgRep, "grey");
                        VoiceAttackPlugin.LogOutput("Trimmed METAR report: " + metarRep, "grey");
                    }
                }
                else
                {
                    Utils.tempoRep = metarRep.Substring(tempoIndex).Trim();
                    Utils.tempoRep = Regex.Replace(Utils.tempoRep, @"\b(TEMPO)\b\s*", "");
                    Utils.tempoRep = Utils.tempoRep.Trim();
                    Utils.tempoRep = "TEMPO " + Utils.tempoRep;
                    metarRep = metarRep.Substring(0, tempoIndex).Trim();
                    metarRep = Regex.Replace(metarRep, @"\b(TEMPO)\b\s*", "");
                    metarRep = metarRep.Trim();
                    Utils.tempoRepExists = true;

                    VoiceAttackPlugin.SetText("TEMPO.Exists", "true");
                    VoiceAttackPlugin.LogOutput("TEMPO section in METAR report detected: " + Utils.tempoRep, "grey");
                    VoiceAttackPlugin.LogOutput("Trimmed METAR report: " + metarRep, "grey");
                }

            }
            else
            {
                if (becmgIndex >= 0)
                {
                    Utils.becmgRep = metarRep.Substring(becmgIndex).Trim();
                    Utils.becmgRep = Regex.Replace(Utils.becmgRep, @"\b(BECMG)\b\s*", "");
                    Utils.becmgRep = Utils.becmgRep.Trim();
                    Utils.becmgRep = "BECMG " + Utils.becmgRep;
                    metarRep = metarRep.Substring(0, becmgIndex).Trim();
                    metarRep = Regex.Replace(metarRep, @"\b(BECMG)\b\s*", "");
                    metarRep = metarRep.Trim();
                    Utils.becmgRepExists = true;

                    VoiceAttackPlugin.SetText("BECMG.Exists", "true");
                    VoiceAttackPlugin.LogOutput("BECMG section in METAR report detected: " + Utils.becmgRep, "grey");
                    VoiceAttackPlugin.LogOutput("Trimmed METAR report: " + metarRep, "grey");
                }
            }

            var components = metarRep.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            if (components.Length == 0)
            {
                VoiceAttackPlugin.LogOutput("No components found in METAR report.", "grey");
            }
            else
            {
                VoiceAttackPlugin.LogOutput(components.Length + " Components found in METAR report", "grey");
                int compIndx = components.Length-1;
                while (compIndx >= 0)
                {
                    VoiceAttackPlugin.LogOutput("Component " + compIndx + ": " + components[compIndx], "grey");
                    compIndx--;
                }
            }

            VoiceAttackPlugin.SetText(reportPrefix + ".Temperature" + reportSufix, "0");
            VoiceAttackPlugin.SetText(reportPrefix + ".DewPoint" + reportSufix, "0");
            VoiceAttackPlugin.SetText(reportPrefix + ".VisibilitySM" + reportSufix, "0");
            VoiceAttackPlugin.SetText(reportPrefix + ".VisibilityKM" + reportSufix, "0");
            VoiceAttackPlugin.SetText(reportPrefix + ".AltimeterA" + reportSufix, "0");
            VoiceAttackPlugin.SetText(reportPrefix + ".AltimeterQ" + reportSufix, "0");
            VoiceAttackPlugin.SetText(reportPrefix + ".WindShear" + reportSufix, "0");
            windType = "0";

            foreach (var component in components)
            {               
                ICAOcheck(component);
                TimeReport(component);
                WindCheck(component);
                WindShearCheck(component);
                TemperatureCheck(component);
                AltimeterCheck(component);
                IntegerForVisibilityCheck(component);
                VisibilityCheck(component);
                SnowtamCheck(component);
                RVRcheck(component);
                WPcheck(component);
                SCcheck(component);
                NosigCheck(component);

                if (Utils.stationIdentified == true)
                {
                    stationIdentified = true;
                }
            }
            
            // ST wrapping up
            if (stExists == true)
            {
                VoiceAttackPlugin.LogOutput("Saving into VA " + stParts.Count + " ST messages", "grey");
                for (int i = 0; i < stParts.Count; i++)
                {
                    VoiceAttackPlugin.SetText(reportPrefix + ".ST" + i + reportSufix, stParts[i]);
                    VoiceAttackPlugin.SetText(reportPrefix + ".STcount" + reportSufix, (stParts.Count-1).ToString());
                    VoiceAttackPlugin.LogOutput("Message " + i + ": " + stParts[i], "grey");
                }
            }

            // RVR wrapping up
            if (rvrExists == true)
            {
                VoiceAttackPlugin.LogOutput("Saving into VA " + rvrParts.Count + " RVR messages", "grey");
                for (int i = 0; i < rvrParts.Count; i++)
                {
                    VoiceAttackPlugin.SetText(reportPrefix + ".RVR" + i + reportSufix, rvrParts[i]);
                    VoiceAttackPlugin.SetText(reportPrefix + ".RVRcount" + reportSufix, (rvrParts.Count - 1).ToString());
                    VoiceAttackPlugin.LogOutput("Message " + i + ": " + rvrParts[i], "grey");
                }
            }

            // WP wrapping up
            if (wpExists == true)
            {
                finalWeatherMessage = wpParts[0];

                VoiceAttackPlugin.LogOutput("Concatenating " + scParts.Count + " WP messages", "grey");
                if (wpParts.Count > 1)
                {
                    finalWeatherMessage = finalWeatherMessage + " with " + wpParts[1];

                    if (wpParts.Count > 2)
                    {
                        for (int i = 2; i < wpParts.Count; i++)
                        {
                            finalWeatherMessage = finalWeatherMessage + " and " + wpParts[i];
                        }
                    }
                }                
                VoiceAttackPlugin.SetText(reportPrefix + ".WP" + reportSufix, finalWeatherMessage);
                VoiceAttackPlugin.LogOutput("WP message:" + finalWeatherMessage, "grey");
            }

            // SC wrapping up
            if (scExists == true)
            {
                VoiceAttackPlugin.LogOutput("Saving into VA " + scParts.Count + " SC messages", "grey");
                for (int i = 0; i < scParts.Count; i++)
                {
                    VoiceAttackPlugin.SetText(reportPrefix + ".SC" + i + reportSufix, scParts[i]);
                    VoiceAttackPlugin.SetText(reportPrefix + ".SCcount" + reportSufix, (scParts.Count-1).ToString());
                    VoiceAttackPlugin.LogOutput("Message " + i + ": " + scParts[i], "grey");
                }
            }

            voiceAttackVariablesSave = false;

        }

        //TAF PARSING
        public void TafReport(string tafRep)
        {
            stationIdentified = false;
            Utils.stationIdentified = false;

            minTemperatureCount = 0;
            maxTemperatureCount = 0;

            string currentSection = "";
            string sectionType = "";
            reportPrefix = "TAF";

            stationIdentifier = "";
            observationTime = "";
            windDirection = "";
            windSpeed = "";
            windGust = "";
            windUnits = "";
            windType = "";
            windFrom = "";
            windTo = "";
            wsalt = "";
            wsdir = "";
            wsspd = "";
            wsunit = "";
            wsmessage = "";
            tfcstMinMessage = "";
            tfcstMaxMessage = "";

            tafSections.Clear();
            interSections.Clear();
            afterInterSections.Clear();

            tafRep = Regex.Replace(tafRep, @"\b(METAR|SPECI|AUTO|TAF)\b\s*", "");
            VoiceAttackPlugin.LogOutput("Removed prefix from TAF: " + tafRep, "grey");

            var tafsections = tafRep.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            VoiceAttackPlugin.LogOutput(tafsections.Length + " Sections found in TAF report.", "grey");            

            sectionIndex = tafsections.Length - 1;
            while (sectionIndex >= 0)
            {
                int interIndex = tafsections[sectionIndex].LastIndexOf("INTER");
                string tafsection = tafsections[sectionIndex].Trim();

                if (interIndex <= 0)
                {                    
                    tafsection = Regex.Replace(tafsection, @"\b(METAR|SPECI|AUTO|TAF)\b\s*", "");
                    VoiceAttackPlugin.LogOutput("Removed prefix from TAF section: " + tafsection, "grey");

                    if (interIndex == 0)
                    {
                        VoiceAttackPlugin.SetText("INTER.Exists" + sectionIndex, "true");
                        VoiceAttackPlugin.LogOutput("INTER section in TAF report detected: " + tafsection, "grey");
                    }

                    tafSections.Add(tafsection);
                    VoiceAttackPlugin.LogOutput("TAF section added into tafSections list: " + tafsection, "grey");

                }
                else
                {
                    VoiceAttackPlugin.LogOutput("INTER sections detected. Rearranging sections", "grey");

                    while (tafsection.Contains("INTER"))
                    {
                        interIndex = tafsection.LastIndexOf("INTER");
                        tafSections.Add(tafsection.Substring(interIndex).Trim());
                        tafsections[sectionIndex] = tafsection.Substring(0, interIndex).Trim();
                        tafsection = tafsections[sectionIndex];
                        VoiceAttackPlugin.LogOutput("INTER subsection found and added into the tafSections list: " + tafsections[sectionIndex], "grey");
                    }
                                        
                    tafSections.Add(tafsections[sectionIndex]);
                    VoiceAttackPlugin.LogOutput("TAF section added into tafSections list: " + tafsection, "grey");

                }
                sectionIndex--;
            }

            tafSections.Reverse();

            string[] newArray = tafSections.ToArray();

            tafsections = newArray;

            VoiceAttackPlugin.SetText(reportPrefix + ".SectionCount", (tafsections.Length - 1).ToString());

            VoiceAttackPlugin.LogOutput("New TAF sections array set", "grey");

            sectionIndex = tafsections.Length - 1;
            while (sectionIndex >= 0)
            {
                VoiceAttackPlugin.LogOutput("Section " + sectionIndex + ": " + tafsections[sectionIndex], "grey");
                sectionIndex--;
            }

            sectionIndex=0;
            currentSection = tafsections[sectionIndex];
            while (sectionIndex < tafsections.Length)
            {                
                reportSufix = sectionIndex.ToString();

                windVariableRange = false;
                windHasGust = false;
                wpExists = false;
                scExists = false;
                rvrExists = false;
                intVisExists = false;

                wpParts.Clear();
                scParts.Clear();
                rvrParts.Clear();
                stParts.Clear();

                int rmkIndex = currentSection.IndexOf("RMK");



                if (rmkIndex >= 0)
                {
                    Utils.rmkRep = currentSection.Substring(rmkIndex).Trim();
                    currentSection = currentSection.Substring(0, rmkIndex).Trim();
                    Utils.rmkRepExists = true;
                    VoiceAttackPlugin.SetText("RMK.Exists", "true");
                    VoiceAttackPlugin.LogOutput("RMK section in TAF report section" + sectionIndex + " detected: " + Utils.rmkRep, "grey");
                    VoiceAttackPlugin.LogOutput("Trimmed TAF section: " + currentSection, "grey");
                }

                currentSection = currentSection.Trim();

                sectionType = "REGULAR";
                if (currentSection.StartsWith("TEMPO"))
                {
                    VoiceAttackPlugin.LogOutput("TAF section is TEMPO", "grey");
                    sectionType = "TEMPO";
                }
                else if (currentSection.StartsWith("BECMG"))
                {
                    VoiceAttackPlugin.LogOutput("TAF section is BECMG", "grey");
                    sectionType = "BECMG";
                }
                else if (currentSection.StartsWith("INTER"))
                {
                    VoiceAttackPlugin.LogOutput("TAF section is INTER", "grey");
                    sectionType = "INTER";
                }
                VoiceAttackPlugin.SetText("TAF.SectionType" + reportSufix, sectionType);                                

                var components = currentSection.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                if (components.Length == 0)
                {
                    VoiceAttackPlugin.LogOutput("No components found in TAF section.", "grey");
                }
                else
                {
                    VoiceAttackPlugin.LogOutput(components.Length + " Components found in TAF section " + sectionIndex, "grey");
                    int compIndx = components.Length - 1;
                    while (compIndx >= 0)
                    {
                        VoiceAttackPlugin.LogOutput("Component " + compIndx + ": " + components[compIndx], "grey");
                        if (components.Length == 1 && (components[compIndx].StartsWith("PROB")))
                        {
                            VoiceAttackPlugin.LogOutput("Alone PROB component", "grey");
                            string nextSection = tafsections[sectionIndex + 1];
                            nextSection = nextSection + " " + components[compIndx];
                            tafsections[sectionIndex+1] = nextSection;
                            components[compIndx] = "";
                        }
                        compIndx--;
                    }
                }

                VoiceAttackPlugin.LogOutput("Rebooting variables.", "grey");
                VoiceAttackPlugin.SetText(reportPrefix + ".Temperature" + reportSufix, "0");
                VoiceAttackPlugin.SetText(reportPrefix + ".DewPoint" + reportSufix, "0");
                VoiceAttackPlugin.SetText(reportPrefix + ".VisibilitySM" + reportSufix, "0");
                VoiceAttackPlugin.SetText(reportPrefix + ".VisibilityKM" + reportSufix, "0");
                VoiceAttackPlugin.SetText(reportPrefix + ".AltimeterA" + reportSufix, "0");
                VoiceAttackPlugin.SetText(reportPrefix + ".AltimeterQ" + reportSufix, "0");
                VoiceAttackPlugin.SetText(reportPrefix + ".ForecastType", "0");
                VoiceAttackPlugin.SetText("TAF.PROBexists" + reportSufix, "false");
                VoiceAttackPlugin.SetText(reportPrefix + ".Empty" + reportSufix, "false");
                VoiceAttackPlugin.SetText("TAF.PROB" + reportSufix, "0");
                VoiceAttackPlugin.SetText(reportPrefix + ".Wind" + reportSufix, "0");
                VoiceAttackPlugin.SetText(reportPrefix + ".WindShear" + reportSufix, "0");
                VoiceAttackPlugin.SetText(reportPrefix + ".TempForecastMin" + reportSufix, "0");
                VoiceAttackPlugin.SetText(reportPrefix + ".TempForecastMax" + reportSufix, "0");
                
                windType = "0";
                VoiceAttackPlugin.LogOutput("Done.", "grey");

                foreach (var component in components)
                {                    
                    if (sectionIndex == 0)
                    {
                        ICAOcheck(component);
                        TimeReport(component);
                    }
                    ForecastPeriod(component);
                    ProbablilityCheck(component);   
                    WindCheck(component);
                    WindShearCheck(component);
                    TempForecastCheck(component);
                    TemperatureCheck(component);
                    AltimeterCheck(component);
                    IntegerForVisibilityCheck(component);
                    VisibilityCheck(component);
                    RVRcheck(component);
                    WPcheck(component);
                    SCcheck(component);
                    VoidCheck(component);

                    if (Utils.stationIdentified == true)
                    {
                        stationIdentified = true;
                    }
                }

                // RVR wrapping up
                if (rvrExists == true)
                {
                    VoiceAttackPlugin.LogOutput("Saving into VA " + rvrParts.Count + " RVR messages for TAF section " + reportSufix, "grey");
                    for (int i = 0; i < rvrParts.Count; i++)
                    {
                        VoiceAttackPlugin.SetText(reportPrefix + ".RVR" + i + reportSufix, rvrParts[i]);
                        VoiceAttackPlugin.SetText(reportPrefix + ".RVRcount" + reportSufix, (rvrParts.Count - 1).ToString());
                        VoiceAttackPlugin.LogOutput("RVR Message " + i + " for TAF section " + reportSufix + ": " + rvrParts[i], "grey");
                    }
                }

                // WP wrapping up
                if (wpExists == true)
                {
                    finalWeatherMessage = wpParts[0];

                    VoiceAttackPlugin.LogOutput("Concatenating " + scParts.Count + " WP messages for TAF section " + reportSufix, "grey");
                    if (wpParts.Count > 1)
                    {
                        finalWeatherMessage = finalWeatherMessage + " with " + wpParts[1];

                        if (wpParts.Count > 2)
                        {
                            for (int i = 2; i < wpParts.Count; i++)
                            {
                                finalWeatherMessage = finalWeatherMessage + " and " + wpParts[i];
                            }
                        }
                    }
                    VoiceAttackPlugin.SetText(reportPrefix + ".WP" + reportSufix, finalWeatherMessage);
                    VoiceAttackPlugin.LogOutput("WP message for TAF section " + reportSufix + ": " + finalWeatherMessage, "grey");
                }

                // SC wrapping up
                if (scExists == true)
                {
                    VoiceAttackPlugin.LogOutput("Saving into VA " + scParts.Count + " SC messages for TAF section " + reportSufix, "grey");
                    for (int i = 0; i < scParts.Count; i++)
                    {
                        VoiceAttackPlugin.SetText(reportPrefix + ".SC" + i + reportSufix, scParts[i]);
                        VoiceAttackPlugin.SetText(reportPrefix + ".SCcount" + reportSufix, (scParts.Count - 1).ToString());
                        VoiceAttackPlugin.LogOutput("Message " + i + " for TAF section " + reportSufix + ": " + scParts[i], "grey");
                    }
                }

                sectionIndex++;
                currentSection = tafsections[sectionIndex];
            }
                                   


        }

        //METAR FORECAST PARSING
        public void AltSection(string altSection)
        {
            reportSufix = "";

            windVariableRange = false;
            windHasGust = false;
            wpExists = false;
            scExists = false;
            rvrExists = false;
            intVisExists = false;

            altSection = altSection.Trim();
            
            if (altSection.StartsWith("RMK"))
            {
                VoiceAttackPlugin.LogOutput("RMK section parsing: " + Utils.rmkRep, "grey");
                reportPrefix = "RMK";                
            }
            if (altSection.StartsWith("TEMPO"))
            {
                VoiceAttackPlugin.LogOutput("TEMPO section parsing: " + Utils.tempoRep, "grey");
                reportPrefix = "TEMPO";                
            }
            if (altSection.StartsWith("BECMG"))
            {
                VoiceAttackPlugin.LogOutput("BECMG section parsing: " + Utils.becmgRep, "grey");
                reportPrefix = "BECMG";
            }

            var components = altSection.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            if (components.Length == 0)
            {
                VoiceAttackPlugin.LogOutput("No components found in sub METAR report.", "grey");
            }
            else
            {
                VoiceAttackPlugin.LogOutput(components.Length + " Components found in sub METAR report", "grey");
                int compIndx = components.Length - 1;
                while (compIndx >= 0)
                {
                    VoiceAttackPlugin.LogOutput("Component " + compIndx + ": " + components[compIndx], "grey");
                    compIndx--;
                }
            }

            wpParts.Clear();
            scParts.Clear();
            rvrParts.Clear();
            stParts.Clear();

            windType = "0";

            VoiceAttackPlugin.SetText(reportPrefix + ".Temperature" + reportSufix, "0");
            VoiceAttackPlugin.SetText(reportPrefix + ".DewPoint" + reportSufix, "0");
            VoiceAttackPlugin.SetText(reportPrefix + ".VisibilitySM" + reportSufix, "0");
            VoiceAttackPlugin.SetText(reportPrefix + ".VisibilityKM" + reportSufix, "0");
            VoiceAttackPlugin.SetText(reportPrefix + ".AltimeterA" + reportSufix, "0");
            VoiceAttackPlugin.SetText(reportPrefix + ".AltimeterQ" + reportSufix, "0");
            VoiceAttackPlugin.SetText(reportPrefix + ".ForecastType", "0");
            VoiceAttackPlugin.SetText(reportPrefix + ".FPend", "0");
            VoiceAttackPlugin.SetText(reportPrefix + ".Wind" + reportSufix, "0");
            VoiceAttackPlugin.SetText(reportPrefix + ".WindShear" + reportSufix, "0");

            windType = "0";

            foreach (var component in components)
            {
                stationIdentified = true;

                ForecastPeriodEnd(component);
                ForecastPeriod(component);
                WindCheck(component);
                TemperatureCheck(component);
                AltimeterCheck(component);
                IntegerForVisibilityCheck(component);
                VisibilityCheck(component);
                SnowtamCheck(component);
                RVRcheck(component);
                WPcheck(component);
                SCcheck(component);
                NosigCheck(component);

            }

            // ST wrapping up
            if (stExists == true)
            {
                VoiceAttackPlugin.LogOutput("Saving into VA " + stParts.Count + " ST messages", "grey");
                for (int i = 0; i < stParts.Count; i++)
                {
                    VoiceAttackPlugin.SetText(reportPrefix + ".ST" + i, stParts[i]);
                    VoiceAttackPlugin.SetText(reportPrefix + ".STcount", (stParts.Count - 1).ToString());
                    VoiceAttackPlugin.LogOutput("Message " + i + ": " + stParts[i], "grey");
                }
            }

            // RVR wrapping up
            if (rvrExists == true)
            {
                VoiceAttackPlugin.LogOutput("Saving into VA " + rvrParts.Count + " RVR messages", "grey");
                for (int i = 0; i < rvrParts.Count; i++)
                {
                    VoiceAttackPlugin.SetText(reportPrefix + ".RVR" + i, rvrParts[i]);
                    VoiceAttackPlugin.SetText(reportPrefix + ".RVRcount", (rvrParts.Count - 1).ToString());
                    VoiceAttackPlugin.LogOutput("Message " + i + ": " + rvrParts[i], "grey");
                }
            }

            // WP wrapping up
            if (wpExists == true)
            {
                finalWeatherMessage = wpParts[0];

                VoiceAttackPlugin.LogOutput("Concatenating " + scParts.Count + " WP messages", "grey");
                if (wpParts.Count > 1)
                {
                    finalWeatherMessage = finalWeatherMessage + " with " + wpParts[1];

                    if (wpParts.Count > 2)
                    {
                        for (int i = 2; i < wpParts.Count; i++)
                        {
                            finalWeatherMessage = finalWeatherMessage + " and " + wpParts[i];
                        }
                    }
                }
                VoiceAttackPlugin.SetText(reportPrefix + ".WP", finalWeatherMessage);
                VoiceAttackPlugin.LogOutput("WP message:" + finalWeatherMessage, "grey");
            }

            // SC wrapping up
            if (scExists == true)
            {
                VoiceAttackPlugin.LogOutput("Saving into VA " + scParts.Count + " SC messages", "grey");
                for (int i = 0; i < scParts.Count; i++)
                {
                    VoiceAttackPlugin.SetText(reportPrefix + ".SC" + i, scParts[i]);
                    VoiceAttackPlugin.SetText(reportPrefix + ".SCcount", (scParts.Count - 1).ToString());
                    VoiceAttackPlugin.LogOutput("Message " + i + ": " + scParts[i], "grey");
                    VoiceAttackPlugin.LogOutput("Var: " + reportPrefix + ".SC" + i, "grey");
                }
            }

        }



        // INDEPENDENT CHECK FUNCTIONS

        public void VoidCheck(string component)
        {
            if (component == "")
            {
                VoiceAttackPlugin.LogOutput("Empty component identified", "grey");
                VoiceAttackPlugin.SetText(reportPrefix + ".Empty" + reportSufix, "true");
            }
        }

        // ICAO DETECTION
        public void ICAOcheck(string component)
        {
            if (Regex.IsMatch(component, @"^[A-Z]{4}$") && stationIdentified == false)
            {                
                VoiceAttackPlugin.LogOutput("Report ICAO component identified: " + component, "grey");
                
                Utils.stationIdentified = true;
                stationIdentifier = component;

                VoiceAttackPlugin.SetText(reportPrefix + ".ICAO", stationIdentifier);
            }
        }

        // TIME OF REPORT DETECTION
        public void TimeReport(string component)
        {
            if (Regex.IsMatch(component, @"^\d{6}Z$"))
            {
                VoiceAttackPlugin.LogOutput("Report TIME component identified: " + component, "grey");

                string day = component.Substring(0, 2);      
                string hour = component.Substring(2, 2);     
                string minute = component.Substring(4, 2);   

                observationTime = $"{hour}:{minute} UTC of day {day}";

                VoiceAttackPlugin.SetText(reportPrefix + ".Time", observationTime);                
            }
        }


        // TEMPERATURE FORECAST DETECTION
        public void TempForecastCheck(string component)
        {
            if ((component.StartsWith("TN")&&component.Contains("/")|| (component.StartsWith("TX") && component.Contains("/"))))
            {                
                VoiceAttackPlugin.LogOutput("Report TEMP FORECAST component identified: " + component, "grey");

                string tfcstType;

                if (component.StartsWith("TN"))
                {
                    minTemperatureCount++;
                    tfcstType = "minimum";
                    string temp = "";
                    string day = "";
                    string hour = "";

                    if (component.Substring(2, 1) == "M")
                    {
                        if (component.Substring(3, 2) == "00")
                        {
                            temp = "0";
                            day = component.Substring(6, 2);
                            hour = component.Substring(8, 2);
                        }
                        else
                        {
                            temp = component.Substring(3, 2);
                            if (temp.StartsWith("0"))
                                temp = temp.Substring(1, 1);
                            temp = "-" + temp;
                            day = component.Substring(6, 2);
                            hour = component.Substring(8, 2);
                        }
                    }
                    else
                    {
                        if (component.Substring(2, 2) == "00")
                        {
                            temp = "0";
                            day = component.Substring(5, 2);
                            hour = component.Substring(7, 2);
                        }
                        else
                        {
                            temp = component.Substring(2, 2);
                            if (temp.StartsWith("0"))
                                temp = temp.Substring(1, 1);
                            day = component.Substring(5, 2);
                            hour = component.Substring(7, 2);
                        }                        
                    }                                     
                    
                    if (tfcstMinMessage == "")
                    {
                        tfcstMinMessage = temp + "ºC at " + hour + ":00 UTC of day " + day;                        
                    }
                    else
                    {
                        tfcstMinMessage = tfcstMinMessage + " and " + temp + "ºC at " + hour + ":00 UTC of day " + day;
                    }
                    VoiceAttackPlugin.SetText(reportPrefix + ".TempForecastMin" + reportSufix, tfcstMinMessage);
                    VoiceAttackPlugin.LogOutput("Temp Forecast message: " + tfcstMinMessage, "grey");

                }
                else if (component.StartsWith("TX"))
                {
                    maxTemperatureCount++;

                    tfcstType = "maximum";
                    string temp = component.Substring(2, 2);
                    if (temp.StartsWith("0"))
                        temp = temp.Substring(1, 1);
                    string day = component.Substring(5, 2);
                    string hour = component.Substring(7, 2);

                    if (tfcstMaxMessage == "")
                    {
                        tfcstMaxMessage = temp + "ºC at " + hour + ":00 UTC of day " + day;
                    }
                    else
                    {
                        tfcstMaxMessage = tfcstMaxMessage + " and " + temp + "ºC at " + hour + ":00 UTC of day " + day;
                    }
                    
                    VoiceAttackPlugin.SetText(reportPrefix + ".TempForecastMax" + reportSufix, tfcstMaxMessage);
                    VoiceAttackPlugin.LogOutput("Temp Forecast message: " + tfcstMaxMessage, "grey");
                }                
            }
        }

        // FORECAST PERIOD DETECTION
        public void ForecastPeriod(string component)
        {
            if (Regex.IsMatch(component, @"\b\d{4}/\d{4}\b|FM\d{6}|FM\d{4}"))
            {   
                if (component.StartsWith("FM"))
                {
                    if (component.Length == 8)
                    {
                        VoiceAttackPlugin.LogOutput("Report RAPID CHANGE PERIOD component identified: " + component, "grey");
                        VoiceAttackPlugin.SetText(reportPrefix + ".ForecastType" + reportSufix, "RAPID");
                        string rapidPeriod = component.Substring(2);
                        string day = rapidPeriod.Substring(0, 2);
                        string hour = rapidPeriod.Substring(2, 2);
                        string minute = rapidPeriod.Substring(4, 2);

                        VoiceAttackPlugin.SetText("TAF.FPstart" + reportSufix, hour + ":" + minute + " UTC of day " + day);
                        VoiceAttackPlugin.LogOutput("Report RAPID CHANGE PERIOD: From " + hour + ":" + minute + " UTC of day " + day, "grey");
                    }
                    else if (component.Length == 6)
                    {
                        VoiceAttackPlugin.LogOutput("Report RAPID CHANGE PERIOD component identified: " + component, "grey");
                        VoiceAttackPlugin.SetText(reportPrefix + ".ForecastType" + reportSufix, "RAPID");
                        string rapidPeriod = component.Substring(2);
                        string hour = rapidPeriod.Substring(0, 2);
                        string minute = rapidPeriod.Substring(2, 2);

                        VoiceAttackPlugin.SetText(reportPrefix + ".FPstart" + reportSufix, hour + ":" + minute + " UTC");
                        VoiceAttackPlugin.LogOutput("Report RAPID CHANGE PERIOD: From " + hour + ":" + minute + " UTC", "grey");
                    }

                }
                else
                {
                    VoiceAttackPlugin.LogOutput("Report FORECAST PERIOD component identified: " + component, "grey");
                    VoiceAttackPlugin.SetText(reportPrefix + ".ForecastType" + reportSufix, "NOTRAPID");

                    var match = Regex.Match(component, @"^(\d{4})/(\d{4})$");
                    string FPstart = match.Groups[1].Value;
                    string FPend = match.Groups[2].Value;

                    string start = FPstart.Substring(2, 2) + ":00 UTC of day " + FPstart.Substring(0, 2);
                    string end = FPend.Substring(2, 2) + ":00 UTC of day " + FPend.Substring(0, 2);

                    VoiceAttackPlugin.SetText("TAF.FPstart" + reportSufix, start);
                    VoiceAttackPlugin.SetText("TAF.FPend" + reportSufix, end);

                    VoiceAttackPlugin.LogOutput("Report FORECAST PERIOD: From " + start + " to " + end, "grey");
                }

            }
        }
        //FORECAST END CHECK
        public void ForecastPeriodEnd(string component)
        {
            if (Regex.IsMatch(component, @"TL\d{4}"))
            {
                VoiceAttackPlugin.LogOutput("Report PERIOD END component identified: " + component, "grey");
                string rapidPeriod = component.Substring(2);
                string hour = rapidPeriod.Substring(0, 2);
                string minute = rapidPeriod.Substring(2, 2);

                VoiceAttackPlugin.SetText(reportPrefix + ".FPend" + reportSufix, hour + ":" + minute + " UTC");
                VoiceAttackPlugin.LogOutput("Report PERIOD END: Until " + hour + ":" + minute + " UTC", "grey");

            }
        }

        // PROBABILITY DETECTION
        public void ProbablilityCheck(string component)
        {
            if (component.StartsWith("PROB"))
            {
                VoiceAttackPlugin.LogOutput("Report PROBABILITY component identified: " + component, "grey");
                string prob = component.Substring(4);
                VoiceAttackPlugin.SetText("TAF.PROB" + reportSufix, prob + "%");
                VoiceAttackPlugin.SetText("TAF.PROBexists" + reportSufix, "true");

                VoiceAttackPlugin.LogOutput("Report PROBABILITY: " + prob + "%", "grey");
            }
        }

        // WIND DETECTION
        public void WindCheck(string component)
        {            
            string windMessage = "0";
            

            if (Regex.IsMatch(component, @"^(?!00000(KT|KMH|MPS)$)(VRB|\d{3})(\d{2})(G\d{2})?(KT|KMH|MPS)$"))
            {
                // WIND REGULAR DETECTION

                windType = "Regular";
                windDirection = "0";
                windSpeed = "0";
                windGust = "0";

                VoiceAttackPlugin.LogOutput("Report WIND component identified: " + component, "grey");

                var match = Regex.Match(component, @"^(VRB|\d{3})(\d{2})(G\d{2})?(KT|KMH|MPS)$");

                windDirection = match.Groups[1].Value;   // "VRB" or three-digit direction
                windSpeed = match.Groups[2].Value;       // Two-digit speed
                windGust = match.Groups[3].Success ? match.Groups[3].Value.Substring(1) : ""; // Optional gust
                windUnits = match.Groups[4].Value;

                VoiceAttackPlugin.SetText(reportPrefix + ".WindType" + reportSufix, "Normal");

                VoiceAttackPlugin.LogOutput("Report WIND component direction: " + windDirection, "grey");
                VoiceAttackPlugin.LogOutput("Report WIND component speed: " + windSpeed, "grey");
                VoiceAttackPlugin.LogOutput("Report WIND component gust: " + windGust, "grey");
                VoiceAttackPlugin.LogOutput("Report WIND component units: " + windUnits, "grey");

                if (!string.IsNullOrEmpty(windGust) && windGust != "0")
                {

                    windHasGust = true;
                }
                    
                // Handle of wind speed units
                if (windUnits == "KMH")
                {
                    VoiceAttackPlugin.LogOutput("Report WIND unit conversion from KMH", "grey");
                    double spdValue = double.Parse(windSpeed);
                    spdValue = spdValue * 0.54;
                    windSpeed = spdValue.ToString("F0");
                    double spdValueG = double.Parse(windGust);
                    spdValueG = spdValueG * 0.54;
                    windGust = spdValueG.ToString("F0");
                }
                if (windUnits == "MPS")
                {
                    VoiceAttackPlugin.LogOutput("Report WIND unit conversion from MPS", "grey");
                    if (windSpeed == "")
                        windSpeed = "0";
                    if (windGust == "")
                        windGust = "0";
                    double spdValue = double.Parse(windSpeed);
                    spdValue = spdValue * 1.94384;
                    windSpeed = spdValue.ToString("F0");
                    double spdValueG = double.Parse(windGust);
                    spdValueG = spdValueG * 1.94384;
                    windGust = spdValueG.ToString("F0");
                }

                // Handle VRB and convert it for readability
                if (windDirection == "VRB")
                {
                    VoiceAttackPlugin.LogOutput("Report WIND VRB detected", "grey");
                    windType = "Variable";
                    windDirection = "Variable";                    
                }
            }

            else if (Regex.IsMatch(component, @"^00000(KT|KMH|MPS)$"))
            {
                // WIND CALM DETECTION

                VoiceAttackPlugin.LogOutput("Report WIND component identified: " + component, "grey");

                windType = "Calm";
                windGust = "0";
                windDirection = "0";
                windSpeed = "0";
            }

            else if (Regex.IsMatch(component, @"^/////(KT|KMH|MPS)$"))
            {
                // NO MEASUREMENT DETECTION

                VoiceAttackPlugin.LogOutput("Report WIND component identified: " + component, "grey");

                windType = "None";
                windGust = "0";
                windDirection = "0";
                windSpeed = "0";
            }

            if (Regex.IsMatch(component, @"^\d{3}V\d{3}$"))
            {
                // WIND VARIABLE DIRECTION DETECTION

                VoiceAttackPlugin.LogOutput("Report WIND component identified: " + component, "grey");

                var match = Regex.Match(component, @"^(\d{3})V(\d{3})$");
                string variableWindFrom = match.Groups[1].Value;  // Start of variable range (e.g., "020")
                string variableWindTo = match.Groups[2].Value;    // End of variable range (e.g., "080")

                windVariableRange = true;

                windFrom = variableWindFrom;
                windTo = variableWindTo;
            }

            if (windType == "Variable")
            {
                if (windHasGust == true)
                {
                    if (windVariableRange == true)
                    {
                        windMessage = "Variable from " + windFrom + "º to " + windTo + "º at " + windSpeed + " knots, gusting to " + windGust + " knots";
                    }
                    else
                    {
                        windMessage = "Variable at " + windSpeed + " knots, gusting to " + windGust + " knots";
                    }                    
                }
                else
                {
                    if (windVariableRange == true)
                    {
                        windMessage = "Variable from " + windFrom + "º to " + windTo + "º at " + windSpeed + " knots";
                    }
                    else
                    {
                        windMessage = "Variable at " + windSpeed + " knots";
                    }
                }
                VoiceAttackPlugin.SetText(reportPrefix + ".Wind" + reportSufix, windMessage);
                VoiceAttackPlugin.SetInt(reportPrefix + ".WindDir" + reportSufix, -1);
                VoiceAttackPlugin.SetInt(reportPrefix + ".WindSpd" + reportSufix, -1);
            }
            else if (windType == "Calm")
            {
                windMessage = windType;
                VoiceAttackPlugin.SetText(reportPrefix + ".Wind" + reportSufix, windMessage);
                VoiceAttackPlugin.SetInt(reportPrefix + ".WindDir" + reportSufix, 0);
                VoiceAttackPlugin.SetInt(reportPrefix + ".WindSpd" + reportSufix, 0);
            }
            else if (windType == "None")
            {
                windMessage = "No wind measurement recorded";
                VoiceAttackPlugin.SetText(reportPrefix + ".Wind" + reportSufix, windMessage);
                VoiceAttackPlugin.SetInt(reportPrefix + ".WindDir" + reportSufix, -1);
                VoiceAttackPlugin.SetInt(reportPrefix + ".WindSpd" + reportSufix, -1);
            }
            else if (windType == "Regular")
            {
                if (windHasGust == true)
                {
                    if (windVariableRange == true)
                    {
                        windMessage = "From " + windDirection + "º at " + windSpeed + " knots, variable from " + windFrom + "º to " + windTo + "º at " + windSpeed + " knots, gusting to " + windGust + " knots";
                    }
                    else
                    {
                        windMessage = "From " + windDirection + "º at " + windSpeed + " knots, gusting to " + windGust + " knots";
                    }
                }
                else
                {
                    if (windVariableRange == true)
                    {
                        windMessage = "From " + windDirection + "º at " + windSpeed + " knots, variable from " + windFrom + "º to " + windTo + "º";
                    }
                    else
                    {
                        windMessage = "From " + windDirection + "º at " + windSpeed + " knots";
                    }
                }
                VoiceAttackPlugin.SetText(reportPrefix + ".Wind" + reportSufix, windMessage);
                VoiceAttackPlugin.SetInt(reportPrefix + ".WindDir" + reportSufix, int.Parse(windDirection));
                VoiceAttackPlugin.SetInt(reportPrefix + ".WindSpd" + reportSufix, int.Parse(windSpeed));
            }            
        }

        // WIND SHEAR DETECTION
        public void WindShearCheck(string component)
        {
            if (Regex.IsMatch(component, @"^WS\d{3}"))
            {
                VoiceAttackPlugin.LogOutput("Report WIND SHEAR component identified: " + component, "grey");

                wsalt = "0";
                wsdir = "0";
                wsspd = "0";
                wsunit = "0";
                wsmessage = "0";

                var match = Regex.Match(component, @"^WS(\d{3})/(\d{3})(\d{2})(KT|KMH|MPS)$");
                if (match.Success)
                {
                    wsalt = match.Groups[1].Value;   // "005" (Shear altitude in hundreds of feet)
                    wsdir = match.Groups[2].Value; // "270" (Wind direction in degrees)
                    wsspd = match.Groups[3].Value;     // "25"  (Wind speed in knots)
                    wsunit = match.Groups[4].Value;     // "KT"  (Wind speed units)
                }

                if (wsalt != "0")
                {
                    if (wsalt.StartsWith("0"))
                        wsalt = wsalt.Substring(1, 2);
                    if (wsalt.StartsWith("0"))
                        wsalt = wsalt.Substring(1, 1);
                    wsalt = wsalt + "00 feet AGL";
                }

                if (wsunit == "KMH")
                {
                    VoiceAttackPlugin.LogOutput("Report WIND SHEAR unit conversion from KMH", "grey");
                    double spdValue = double.Parse(wsspd);
                    spdValue = spdValue * 0.54;
                    wsspd = spdValue.ToString("F0");
                }
                if (wsunit == "MPS")
                {
                    VoiceAttackPlugin.LogOutput("Report WIND SHEAR unit conversion from MPS", "grey");
                    if (wsspd == "")
                        wsspd = "0";
                    double spdValue = double.Parse(wsspd);
                    spdValue = spdValue * 1.94384;
                    wsspd = spdValue.ToString("F0");
                }

                wsmessage = wsalt + " - From " + wsdir + "º at " + wsspd + " knots";

                VoiceAttackPlugin.SetText(reportPrefix + ".WindShear" + reportSufix, wsmessage);

            }
        }

        // TEMPERATURE DETECTION
        public void TemperatureCheck(string component)
        {        
            if (Regex.IsMatch(component, @"^(M?\d{2})/(M?\d{2})$"))
            {               
                VoiceAttackPlugin.LogOutput("Report TEMP component identified: " + component, "grey");

                string temperature = string.Empty;
                string dewPoint = string.Empty;

                var parts = component.Split('/'); // Split on the '/' character

                temperature = parts[0];
                dewPoint = parts[1];

                VoiceAttackPlugin.LogOutput("Report TEMP component Temp: " + temperature, "grey");
                VoiceAttackPlugin.LogOutput("Report TEMP component Dew: " + dewPoint, "grey");

                if (temperature.StartsWith("M"))
                {
                    if (temperature == "M00")
                    {
                        temperature = "00";
                    }
                    else
                    {
                        temperature = temperature.Substring(1); // Remove the 'M'
                        if (temperature.StartsWith("0"))
                            temperature = temperature.Substring(1, 1);
                        temperature = "-" + temperature; // Prepend negative sign to the number
                    }
                    VoiceAttackPlugin.LogOutput("Report TEMP component Temp Mod: " + dewPoint, "grey");
                }

                if (dewPoint.StartsWith("M"))
                {
                    if (dewPoint == "M00")
                    {
                        dewPoint = "00";
                    }
                    else
                    {
                        dewPoint = dewPoint.Substring(1); // Remove the 'M'
                        if (dewPoint.StartsWith("0"))
                            dewPoint = dewPoint.Substring(1, 1);
                        dewPoint = "-" + dewPoint; // Prepend negative sign to the number
                    }
                    VoiceAttackPlugin.LogOutput("Report TEMP component Dew Mod: " + dewPoint, "grey");
                }
                if (dewPoint.StartsWith("0"))
                    dewPoint = dewPoint.Substring(1, 1);
                if (temperature.StartsWith("0"))
                    temperature = temperature.Substring(1, 1);

                decimal decTemperature = decimal.Parse(temperature);

                VoiceAttackPlugin.SetDecimal(reportPrefix + ".DecTemperature" + reportSufix, decTemperature);
                VoiceAttackPlugin.SetText(reportPrefix + ".Temperature" + reportSufix, temperature + "ºC");              
                VoiceAttackPlugin.SetText(reportPrefix + ".DewPoint" + reportSufix, dewPoint + "ºC");
            }
        }

        // ALTIMETER DETECTION
        public void AltimeterCheck(string component)
        {
            if (Regex.IsMatch(component, @"^(A|Q)\d{4}$"))
            {               
                VoiceAttackPlugin.LogOutput("Report QNH component identified: " + component, "grey");

                string altimeterUnit = component.Substring(0, 1);
                string altimeterValue = component.Substring(1, 4);
                VoiceAttackPlugin.LogOutput("QNH component value: " + altimeterValue, "grey");
                VoiceAttackPlugin.LogOutput("QNH component unit: " + altimeterUnit, "grey");

                if (altimeterUnit == "A")
                {
                    int altimeterValInt = int.Parse(altimeterValue);
                    string altimeterINHG = altimeterValue.Substring(0, 2) + "." + altimeterValue.Substring(2, 2);
                    double altValue = double.Parse(altimeterValue);
                    altValue = altValue * 33.8639 / 100;
                    string altimeterHPA = altValue.ToString("F0");
                    VoiceAttackPlugin.LogOutput("QNH component formatted INHG: " + altimeterINHG, "grey");
                    VoiceAttackPlugin.LogOutput("QNH component formatted HPA: " + altimeterHPA, "grey");

                    VoiceAttackPlugin.SetInt(reportPrefix + ".Altimeter" + reportSufix, altimeterValInt);
                    VoiceAttackPlugin.SetText(reportPrefix + ".AltimeterA" + reportSufix, altimeterINHG + "inHg");
                    VoiceAttackPlugin.SetText(reportPrefix + ".AltimeterQ" + reportSufix, altimeterHPA + "hPa");                    
                }
                if (altimeterUnit == "Q")
                {
                    string altimeterHPA = altimeterValue;
                    double altValue = double.Parse(altimeterValue);
                    altValue = altValue * 0.02953;
                    double altValue2 = altValue * 100;
                    int altimeterValInt = (int)altValue2;
                    string altimeterINHG = altValue.ToString("F2");
                    VoiceAttackPlugin.LogOutput("QNH component formatted INHG: " + altimeterINHG, "grey");
                    VoiceAttackPlugin.LogOutput("QNH component formatted HPA: " + altimeterHPA, "grey");

                    VoiceAttackPlugin.SetInt(reportPrefix + ".Altimeter" + reportSufix, altimeterValInt);
                    VoiceAttackPlugin.SetText(reportPrefix + ".AltimeterA" + reportSufix, altimeterINHG + "inHg");
                    VoiceAttackPlugin.SetText(reportPrefix + ".AltimeterQ" + reportSufix, altimeterHPA + "hPa");                                        
                }               
            }
        }

        // VISIBILITY DETECTION
        public void VisibilityCheck(string component)
        {
            if (Regex.IsMatch(component, @"^P?\d{1,3}(SM|KM)$|^\d{1}\/\d{1,2}SM$|^\d{4}$"))
            {
                VoiceAttackPlugin.LogOutput("Report VISIBILITY component identified: " + component, "grey");

                visibilityValueKM = "0";
                visibilityValueSM = "0";
                
                if (component.EndsWith("KM"))
                {
                    if (component.StartsWith("P"))
                    {
                        component = component.Substring(1);
                        VoiceAttackPlugin.LogOutput("Report VISIBILITY in KM with PLUS at beginning", "grey");
                        visibilityValueKM = component.Substring(0, component.Length - 2);
                        double visValue = double.Parse(visibilityValueKM);
                        visValue = visValue * 0.621371;
                        visibilityValueSM = visValue.ToString("F1");

                        VoiceAttackPlugin.SetText(reportPrefix + ".VisibilitySM" + reportSufix, "More than " + visibilityValueSM + " Statute Miles");
                        VoiceAttackPlugin.SetText(reportPrefix + ".VisibilityKM" + reportSufix, "More than " + visibilityValueKM + " Kilometers");
                    }
                    else
                    {
                        VoiceAttackPlugin.LogOutput("Report VISIBILITY in KM", "grey");
                        visibilityValueKM = component.Substring(0, component.Length - 2);
                        double visValue = double.Parse(visibilityValueKM);
                        visValue = visValue * 0.621371;
                        visibilityValueSM = visValue.ToString("F1");

                        VoiceAttackPlugin.SetText(reportPrefix + ".VisibilitySM" + reportSufix, visibilityValueSM + " Statute Miles");
                        VoiceAttackPlugin.SetText(reportPrefix + ".VisibilityKM" + reportSufix, visibilityValueKM + " Kilometers");
                    }                    
                }
                else if (component.EndsWith("SM"))
                {
                    if (component.Contains("/"))
                    {

                        VoiceAttackPlugin.LogOutput("Report VISIBILITY in fractional SM", "grey");
                        visibilityValueSM = component.Substring(0, component.Length - 2);

                        string[] parts = visibilityValueSM.Split(new char[] { '/' }, 2);
                        string num = parts[0];  // Everything before the slash
                        string den = parts[1];   // Everything after the slash

                        double numD = double.Parse(num);
                        double denD = double.Parse(den);
                        double visValue = numD / denD;

                        VoiceAttackPlugin.LogOutput("Report VISIBILITY fractional with numerator " + numD + " and denominator " + denD + " resulting decimal " + visValue, "grey");

                        if (intVisExists == true)
                        {
                            double intVis = intVisValue;

                            VoiceAttackPlugin.LogOutput("Report VISIBILITY integer for fractional detected: " + intVis, "grey");

                            visValue = visValue + intVis;

                            VoiceAttackPlugin.LogOutput("Report VISIBILITY total value: " + visValue, "grey");
                        }

                        visibilityValueSM = visValue.ToString("F1");
                        visValue = visValue * 1.60934;
                        visibilityValueKM = visValue.ToString("F1");

                        VoiceAttackPlugin.SetText(reportPrefix + ".VisibilitySM" + reportSufix, visibilityValueSM + " Statute Miles");
                        VoiceAttackPlugin.SetText(reportPrefix + ".VisibilityKM" + reportSufix, visibilityValueKM + " Kilometers");
                    }
                    else
                    {
                        if (component.StartsWith("P"))
                        {
                            component = component.Substring(1);
                            VoiceAttackPlugin.LogOutput("Report VISIBILITY in SM with PLUS at beginning", "grey");
                            visibilityValueSM = component.Substring(0, component.Length - 2);
                            double visValue = double.Parse(visibilityValueSM);
                            visValue = visValue * 1.60934;
                            visibilityValueKM = visValue.ToString("F1");

                            VoiceAttackPlugin.SetText(reportPrefix + ".VisibilitySM" + reportSufix, "More than " + visibilityValueSM + " Statute Miles");
                            VoiceAttackPlugin.SetText(reportPrefix + ".VisibilityKM" + reportSufix, "More than " + visibilityValueKM + " Kilometers");
                        }
                        else
                        {
                            VoiceAttackPlugin.LogOutput("Report VISIBILITY in SM", "grey");
                            visibilityValueSM = component.Substring(0, component.Length - 2);
                            double visValue = double.Parse(visibilityValueSM);
                            visValue = visValue * 1.60934;
                            visibilityValueKM = visValue.ToString("F1");

                            VoiceAttackPlugin.SetText(reportPrefix + ".VisibilitySM" + reportSufix, visibilityValueSM + " Statute Miles");
                            VoiceAttackPlugin.SetText(reportPrefix + ".VisibilityKM" + reportSufix, visibilityValueKM + " Kilometers");
                        }
                        
                    }

                }
                else if (component == "9999")
                {
                    VoiceAttackPlugin.LogOutput("Report VISIBILITY in 4 numbers", "grey");

                    VoiceAttackPlugin.SetText(reportPrefix + ".VisibilitySM" + reportSufix, "Greater than 6 Statute Miles");
                    VoiceAttackPlugin.SetText(reportPrefix + ".VisibilityKM" + reportSufix, "Greater than 10 Kilometers");
                }
                else
                {
                    VoiceAttackPlugin.LogOutput("Report VISIBILITY in indetermined format", "grey");
                    double visValue = double.Parse(component);
                    visValue = visValue / 1000;
                    visibilityValueKM = visValue.ToString("F1");
                    visValue = visValue * 0.621371;
                    visibilityValueSM = visValue.ToString("F1");

                    VoiceAttackPlugin.SetText(reportPrefix + ".VisibilitySM" + reportSufix, visibilityValueSM + " Statute Miles");
                    VoiceAttackPlugin.SetText(reportPrefix + ".VisibilityKM" + reportSufix, visibilityValueKM + " Kilometers");
                }

            }
        }

        // NOSIG DETECTION
        public void NosigCheck(string component)
        {
            VoiceAttackPlugin.SetText(reportPrefix + ".NOSIG" + reportSufix, "false");

            if (Regex.IsMatch(component, @"^(NOSIG)$"))
            {
                VoiceAttackPlugin.LogOutput("NOSIG component detected", "grey");

                VoiceAttackPlugin.SetText(reportPrefix + ".NOSIG" + reportSufix, "true");

            }
        }

        // VISIBILITY INTEGER DETECTION
        public void IntegerForVisibilityCheck(string component)
        {
            if (Regex.IsMatch(component, @"^\d$"))
            {
                VoiceAttackPlugin.LogOutput("Report component potential VISIBILITY integer detected: " + component, "grey");

                intVisExists = true;
                
                intVisValue = int.Parse(component);
            }

        }

        //SNOWTAM DETECTION
        public void SnowtamCheck(string component)
        {
            VoiceAttackPlugin.SetText(reportPrefix + ".STcount" + reportSufix, "-1");

            if (((Regex.IsMatch(component, @"^R\d{2}")) && (Regex.IsMatch(component, @"\d{6}$"))) || ((Regex.IsMatch(component, @"^R\d{2}")) && component.Contains("CLRD")))
            {
                stExists = true;                

                snowtamIndex++;

                string snowtamMessage = "0";
                string snowtamBraking = "0";

                stRunway = "NA";
                stDeposit = "NA";
                stContamination = "NA";
                stDepth = "NA";
                stBraking = "NA";

                VoiceAttackPlugin.LogOutput("Report SNOWTAM component identified: " + component, "grey");

                string[] parts = component.Split(new char[] { '/' }, 2);

                string runway = parts[0];  // Everything before the slash
                string snowtam = parts[1];   // Everything after the slash
                VoiceAttackPlugin.LogOutput("Report SNOWTAM runway: " + runway + ".", "grey");
                VoiceAttackPlugin.LogOutput("Report SNOWTAM code: " + snowtam + ".", "grey");

                runway = runway.Substring(1);
                VoiceAttackPlugin.LogOutput("Report SNOWTAM runway mod: " + runway + ".", "grey");
                VoiceAttackPlugin.SetText(reportPrefix + ".STRunway" + snowtamIndex, runway);

                if (snowtam.StartsWith("CLRD"))
                {
                    string d = "";
                    if (snowtam.Substring(4, 2) == "91")
                    {
                        d = "Poor";
                    }
                    else if (snowtam.Substring(4, 2) == "92")
                    {
                        d = "Medium-Poor";
                    }
                    else if (snowtam.Substring(4, 2) == "93")
                    {
                        d = "Medium";
                    }
                    else if (snowtam.Substring(4, 2) == "94")
                    {
                        d = "Medium-Good";
                    }
                    else if (snowtam.Substring(4, 2) == "95")
                    {
                        d = "Good";
                    }
                    else if (snowtam.Substring(4, 2) == "99")
                    {
                        d = "Undetermined";
                    }
                    else
                    {
                        d = "0." + snowtam.Substring(4, 2);
                    }

                    VoiceAttackPlugin.LogOutput("Report SNOWTAM braking conditions: " + d + ".", "grey");
                    stBraking = d;

                    if (stBraking.StartsWith("0"))
                    {
                        snowtamBraking = "with a Friction Coefficient of " + stBraking;
                    }
                    else
                    {
                        snowtamBraking = "with " + stBraking + " braking conditions";
                    }
                    snowtamMessage = "Runway cleared " + snowtamBraking;
                    stParts.Add(snowtamMessage);
                }
                else
                {
                    string d = "";
                    switch (snowtam.Substring(0, 1))
                    {
                        case "0":
                            d = "Clear and dry";
                            break;
                        case "1":
                            d = "Damp";
                            break;
                        case "2":
                            d = "Water";
                            break;
                        case "3":
                            d = "Rime or Frost";
                            break;
                        case "4":
                            d = "Dry Snow";
                            break;
                        case "5":
                            d = "Wet Snow";
                            break;
                        case "6":
                            d = "Slush";
                            break;
                        case "7":
                            d = "Ice";
                            break;
                        case "8":
                            d = "Compacted or Rolled Snow";
                            break;
                        case "9":
                            d = "Frozen Ruts or Ridges";
                            break;
                    }
                    VoiceAttackPlugin.LogOutput("Report SNOWTAM deposit: " + d + ".", "grey");
                    stDeposit = d;

                    d = "";
                    switch (snowtam.Substring(1, 1))
                    {
                        case "1":
                            d = "10% or less";
                            break;
                        case "2":
                            d = "11% to 25%";
                            break;
                        case "5":
                            d = "26 to 50%";
                            break;
                        case "9":
                            d = "51 to 100%";
                            break;
                    }
                    VoiceAttackPlugin.LogOutput("Report SNOWTAM contamination: " + d + ".", "grey");
                    stContamination = d;

                    d = "";
                    if (snowtam.Substring(2, 2) == "00")
                    {
                        d = "less than 1mm";
                    }
                    else if (snowtam.Substring(2, 2) == "92")
                    {
                        d = "100 mm";
                    }
                    else if (snowtam.Substring(2, 2) == "93")
                    {
                        d = "150 mm";
                    }
                    else if (snowtam.Substring(2, 2) == "94")
                    {
                        d = "200 mm";
                    }
                    else if (snowtam.Substring(2, 2) == "95")
                    {
                        d = "250 mm";
                    }
                    else if (snowtam.Substring(2, 2) == "96")
                    {
                        d = "300 mm";
                    }
                    else if (snowtam.Substring(2, 2) == "97")
                    {
                        d = "350 mm";
                    }
                    else if (snowtam.Substring(2, 2) == "98")
                    {
                        d = "400 mm";
                    }
                    else if (snowtam.Substring(2, 2) == "99")
                    {
                        d = "NOT OPERATIONAL";
                    }
                    else
                    {
                        d = snowtam.Substring(2, 2) + " mm";
                        if (d.StartsWith("0"))
                            d = d.Substring(1, 4);
                    }

                    VoiceAttackPlugin.LogOutput("Report SNOWTAM depth: " + d + ".", "grey");
                    stDepth = d;

                    d = "";
                    if (snowtam.Substring(4, 2) == "91")
                    {
                        d = "Poor";
                    }
                    else if (snowtam.Substring(4, 2) == "92")
                    {
                        d = "Medium-Poor";
                    }
                    else if (snowtam.Substring(4, 2) == "93")
                    {
                        d = "Medium";
                    }
                    else if (snowtam.Substring(4, 2) == "94")
                    {
                        d = "Medium-Good";
                    }
                    else if (snowtam.Substring(4, 2) == "95")
                    {
                        d = "Good";
                    }
                    else if (snowtam.Substring(4, 2) == "99")
                    {
                        d = "Undetermined";
                    }
                    else
                    {
                        d = "0." + snowtam.Substring(4, 2);
                    }

                    VoiceAttackPlugin.LogOutput("Report SNOWTAM braking conditions: " + d + ".", "grey");
                    stBraking = d;

                    if (stBraking.StartsWith("0"))
                    {
                        snowtamBraking = "with a Friction Coefficient of " + stBraking;
                    }
                    else
                    {
                        snowtamBraking = "with " + stBraking + " braking conditions";
                    }

                    snowtamMessage = stContamination + " covered by " + stDepth + " of " + stDeposit + ", " + snowtamBraking;
                    stParts.Add(snowtamMessage);
                }
                
            }
        }

        // RUNWAY VISUAL RANGE DETECTION
        public void RVRcheck(string component)
        {
            VoiceAttackPlugin.SetText(reportPrefix + ".RVRcount" + reportSufix, "-1");

            if ((Regex.IsMatch(component, @"^R\d{2}")) && (Regex.IsMatch(component, @"^(?!.*\d{6}$).*$")) && !component.Contains("CLRD"))
            {
                VoiceAttackPlugin.LogOutput("Report RVR component identified: " + component, "grey");

                rvrExists = true;

                rvrIndex++;

                string rvrMessage = "0";

                rvrTrend = "0";
                rvrVisibility = "0";
                rvrRunway = "0";
                rvrUnit = "0";

                string[] parts = component.Split(new char[] { '/' }, 2);

                string runway = parts[0];  // Everything before the slash
                string rvrVis = parts[1];   // Everything after the slash
                VoiceAttackPlugin.LogOutput("Report RVR component runway: " + runway + ".", "grey");
                VoiceAttackPlugin.LogOutput("Report RVR component visibility: " + rvrVis + ".", "grey");                           

                runway = runway.Substring(1);
                VoiceAttackPlugin.LogOutput("Report RVR component runway mod: " + runway + ".", "grey");

                VoiceAttackPlugin.SetText(reportPrefix + ".RVRRunway" + reportSufix + rvrIndex, runway);

                int vIndex = rvrVis.IndexOf('V');
                if (vIndex >= 0)  // Check if 'V' is found
                {
                    if (vIndex > 3 && vIndex < rvrVis.Length - 4)
                    {
                        // Grab the previous 4 characters before the 'V'
                        string visFrom = rvrVis.Substring(vIndex - 4, 4);
                        if (visFrom.StartsWith("0"))
                            visFrom = visFrom.Substring(1, 3);

                        // Grab the next 4 characters after the 'V'
                        string visTo = rvrVis.Substring(vIndex + 1, 4);
                        if (visTo.StartsWith("0"))
                            visTo = visTo.Substring(1, 3);

                        VoiceAttackPlugin.LogOutput("Report RVR variability detected: From" + visFrom + "To" + visTo + ".", "grey");

                        rvrVisibility = "Variable from " + visFrom + " to " + visTo;
                    }

                }
                else if (rvrVis.StartsWith("M"))
                {
                    string visDis = rvrVis.Substring(1, 4);
                    if (visDis.StartsWith("0"))
                        visDis = visDis.Substring(1, 3);

                    VoiceAttackPlugin.LogOutput("Report RVR range detected: Less than" + visDis + ".", "grey");

                    rvrVisibility = "Less than " + visDis;
                }

                else if (rvrVisibility.StartsWith("P"))
                {
                    string visDis = rvrVis.Substring(1, 4);
                    if (visDis.StartsWith("0"))
                        visDis = visDis.Substring(1, 3);

                    VoiceAttackPlugin.LogOutput("Report RVR range detected: More than" + visDis + ".", "grey");

                    rvrVisibility = "More than " + visDis;
                }
                else
                {
                    string visDis = rvrVis.Substring(0, 4);
                    if (visDis.StartsWith("0"))
                        visDis = visDis.Substring(1, 3);

                    VoiceAttackPlugin.LogOutput("Report RVR range detected: Equal to " + visDis + ".", "grey");

                    rvrVisibility = visDis;
                }

                if (rvrVis.EndsWith("D"))
                {
                    VoiceAttackPlugin.LogOutput("Report RVR trend detected: Down", "grey");

                    rvrTrend = "Decreasing";                    
                }
                else if (rvrVis.EndsWith("U"))
                {
                    VoiceAttackPlugin.LogOutput("Report RVR trend detected: Up", "grey");

                    rvrTrend = "Increasing";
                }
                else if (rvrVis.EndsWith("N"))
                {
                    VoiceAttackPlugin.LogOutput("Report RVR trend detected: No Change", "grey");

                    rvrTrend = "No change";
                }

                if (rvrVis.Contains("FT"))
                {
                    VoiceAttackPlugin.LogOutput("Report RVR unit detected: Feet", "grey");

                    rvrUnit = " feet";
                }
                else
                {
                    VoiceAttackPlugin.LogOutput("Report RVR unit assumed: Meters", "grey");

                    rvrUnit = " meters";
                }

                if (rvrTrend != "0")
                {
                    rvrMessage = rvrVisibility + rvrUnit + " with " + rvrTrend;
                }               
                else
                {
                    rvrMessage = rvrVisibility + rvrUnit;
                }

                rvrParts.Add(rvrMessage);                
            }
        }

        // WEATHER PHENOMENA DETECTION
        public void WPcheck(string component)
        {
            VoiceAttackPlugin.SetText(reportPrefix + ".WP" + reportSufix, "0");

            string[] wpExcluded = { "RETS", "CAVOK", "NOSIG", "AUTO", "CLR", "NSC", "TEMPO", "NCD", "BECMG" , "SKC" };
            if (Regex.IsMatch(component, @"^[-+A-Z]{2,}$") && !wpExcluded.Contains(component) && stationIdentified == true)
            {
                VoiceAttackPlugin.LogOutput("Report WP component identified: " + component + ".", "grey");

                wpExists = true;
                Utils.wpQuantity++;

                string intensity = "";
                string descriptor = "";
                string wp2 = "";
                int preDescriptor = 1;
                string weatherPhenomenaType = component;

                if (component.StartsWith("-"))
                {
                    intensity = "Light"; // Light intensity
                    weatherPhenomenaType = component.Substring(1); // Remove the first character
                }
                else if (component.StartsWith("+"))
                {
                    intensity = "Heavy"; // Heavy intensity
                    weatherPhenomenaType = component.Substring(1); // Remove the first character
                    if (voiceAttackVariablesSave == true)
                    {
                        VoiceAttackPlugin.SetText("heavyPrecipitation", "true");
                    }
                }

                VoiceAttackPlugin.LogOutput("Report WP component intensity: " + intensity, "grey");                                

                if (weatherPhenomenaType.Length == 6)
                {
                    wp2 = weatherPhenomenaType.Substring(4);
                    weatherPhenomenaType = weatherPhenomenaType.Substring(0,4);
                    VoiceAttackPlugin.LogOutput("Report WP Second weather detected: " + wp2, "grey");
                    VoiceAttackPlugin.LogOutput("Report WP Main: " + weatherPhenomenaType, "grey");
                }

                if (weatherPhenomenaType.Length == 4)
                {
                    descriptor = weatherPhenomenaType.Substring(0, 2);
                    weatherPhenomenaType = weatherPhenomenaType.Substring(2);
                    VoiceAttackPlugin.LogOutput("Report WP descriptor detected: " + descriptor, "grey");
                    VoiceAttackPlugin.LogOutput("Report WP Weather: " + weatherPhenomenaType, "grey");

                    string descriptor2 = descriptor;

                    switch (descriptor2)
                    {
                        case "MI":
                            descriptor = "Shallow";
                            break;
                        case "BC":
                            descriptor = "Patches of";
                            break;
                        case "DR":
                            descriptor = "Low Drifting";
                            break;
                        case "DZ":
                            descriptor = "Drizzle";
                            preDescriptor = -1;
                            if (voiceAttackVariablesSave == true)
                            {
                                VoiceAttackPlugin.SetText("runwayState", "WET");
                                VoiceAttackPlugin.SetText("visibleMoisture", "true");
                            }
                            break;
                        case "BL":
                            descriptor = "Blowing";
                            break;
                        case "SH":
                            descriptor = "Showers";
                            preDescriptor = 0;
                            if (voiceAttackVariablesSave == true)
                            {
                                VoiceAttackPlugin.SetText("visibleMoisture", "true");
                            }
                            break;
                        case "TS":
                            descriptor = "Thunderstorm with";
                            if (voiceAttackVariablesSave == true)
                            {
                                VoiceAttackPlugin.SetText("visibleMoisture", "true");
                                VoiceAttackPlugin.SetText("heavyPrecipitation", "true");
                            }
                            break;
                        case "FZ":
                            descriptor = "Freezing";
                            break;
                        case "PR":
                            descriptor = "Partial";
                            break;
                        case "VC":
                            descriptor = "in the vicinity";
                            preDescriptor = 0;
                            break;
                        case "RA":
                            descriptor = "Rain";
                            preDescriptor = -1;
                            if (voiceAttackVariablesSave == true)
                            {
                                VoiceAttackPlugin.SetText("runwayState", "WET");
                                VoiceAttackPlugin.SetText("visibleMoisture", "true");
                            }
                            
                            break;
                        case "RE":
                            descriptor = "Recent";
                            break;
                        case "SN":
                            descriptor = "Snow";
                            preDescriptor = -1;
                            if (voiceAttackVariablesSave == true)
                            {
                                VoiceAttackPlugin.SetText("runwayState", "WET");
                                VoiceAttackPlugin.SetText("visibleMoisture", "true");
                            }
                            break;
                    }
                }

                string weatherPhenomenaType2 = weatherPhenomenaType;
                switch (weatherPhenomenaType2)
                {
                    case "BR":
                        weatherPhenomenaType = "Mist";
                        if (voiceAttackVariablesSave == true)
                        {
                            VoiceAttackPlugin.SetText("visibleMoisture", "true");
                        }
                        break;
                    case "DS":
                        weatherPhenomenaType = "Dust Storm";
                        break;
                    case "DU":
                        weatherPhenomenaType = "Widespread Dust";
                        break;
                    case "DZ":
                        weatherPhenomenaType = "Drizzle";
                        if (voiceAttackVariablesSave == true)
                        {
                            VoiceAttackPlugin.SetText("runwayState", "WET");
                        }
                        break;
                    case "FG":
                        weatherPhenomenaType = "Fog";
                        if (voiceAttackVariablesSave == true)
                        {
                            VoiceAttackPlugin.SetText("visibleMoisture", "true");
                        }
                        break;
                    case "FC":
                        weatherPhenomenaType = "Funnel Cloud";
                        if (voiceAttackVariablesSave == true)
                        {
                            VoiceAttackPlugin.SetText("visibleMoisture", "true");
                        }
                        break;
                    case "FU":
                        weatherPhenomenaType = "Smoke";
                        break;
                    case "GR":
                        weatherPhenomenaType = "Hail";
                        if (voiceAttackVariablesSave == true)
                        {
                            VoiceAttackPlugin.SetText("runwayState", "WET");
                        }
                        break;
                    case "GS":
                        weatherPhenomenaType = "Small Hail";
                        if (voiceAttackVariablesSave == true)
                        {
                            VoiceAttackPlugin.SetText("runwayState", "WET");
                        }
                        break;
                    case "HZ":
                        weatherPhenomenaType = "Haze";
                        break;
                    case "IC":
                        weatherPhenomenaType = "Ice Crystals";
                        if (voiceAttackVariablesSave == true)
                        {
                            VoiceAttackPlugin.SetText("runwayState", "WET");
                            VoiceAttackPlugin.SetText("visibleMoisture", "true");
                        }
                        break;
                    case "PL":
                        weatherPhenomenaType = "Ice Pellets";
                        if (voiceAttackVariablesSave == true)
                        {
                            VoiceAttackPlugin.SetText("runwayState", "WET");
                            VoiceAttackPlugin.SetText("visibleMoisture", "true");
                        }
                        break;
                    case "PO":
                        weatherPhenomenaType = "Dust/Sand Whirls";
                        break;
                    case "PY":
                        weatherPhenomenaType = "Spray";
                        break;
                    case "RA":
                        weatherPhenomenaType = "Rain";
                        if (voiceAttackVariablesSave == true)
                        {
                            VoiceAttackPlugin.SetText("runwayState", "WET");
                            VoiceAttackPlugin.SetText("visibleMoisture", "true");
                        }
                        break;
                    case "SA":
                        weatherPhenomenaType = "Sand";
                        break;
                    case "SG":
                        weatherPhenomenaType = "Snow Grains";
                        if (voiceAttackVariablesSave == true)
                        {
                            VoiceAttackPlugin.SetText("runwayState", "WET");
                            VoiceAttackPlugin.SetText("visibleMoisture", "true");
                        }
                        break;
                    case "SH":
                        weatherPhenomenaType = "Showers";
                        if (voiceAttackVariablesSave == true)
                        {
                            VoiceAttackPlugin.SetText("visibleMoisture", "true");
                        }
                        break;
                    case "SN":
                        weatherPhenomenaType = "Snow";
                        if (voiceAttackVariablesSave == true)
                        {
                            VoiceAttackPlugin.SetText("runwayState", "WET");
                            VoiceAttackPlugin.SetText("visibleMoisture", "true");
                        }
                        break;
                    case "SQ":
                        weatherPhenomenaType = "Squalls";
                        break;
                    case "SS":
                        weatherPhenomenaType = "Sandstorm";
                        break;
                    case "VA":
                        weatherPhenomenaType = "Volcanic Ash";
                        break;
                    case "UP":
                        weatherPhenomenaType = "Unknown Precipitation";
                        break;
                }

                string weatherMessage = weatherPhenomenaType;

                if (wp2 != "")
                {
                    string weatherPhenomenaType3 = wp2;
                    switch (weatherPhenomenaType3)
                    {
                        case "BR":
                            wp2 = "Mist";
                            if (voiceAttackVariablesSave == true)
                            {
                                VoiceAttackPlugin.SetText("visibleMoisture", "true");
                            }
                            break;
                        case "DS":
                            wp2 = "Dust Storm";
                            break;
                        case "DU":
                            wp2 = "Widespread Dust";
                            break;
                        case "DZ":
                            wp2 = "Drizzle";
                            if (voiceAttackVariablesSave == true)
                            {
                                VoiceAttackPlugin.SetText("runwayState", "WET");
                                VoiceAttackPlugin.SetText("visibleMoisture", "true");
                            }
                            break;
                        case "FG":
                            wp2 = "Fog";
                            if (voiceAttackVariablesSave == true)
                            {
                                VoiceAttackPlugin.SetText("visibleMoisture", "true");
                            }
                            break;
                        case "FC":
                            wp2 = "Funnel Cloud";
                            if (voiceAttackVariablesSave == true)
                            {
                                VoiceAttackPlugin.SetText("visibleMoisture", "true");
                            }
                            break;
                        case "FU":
                            wp2 = "Smoke";
                            break;
                        case "GR":
                            wp2 = "Hail";
                            if (voiceAttackVariablesSave == true)
                            {
                                VoiceAttackPlugin.SetText("runwayState", "WET");
                                VoiceAttackPlugin.SetText("visibleMoisture", "true");
                            }
                            break;
                        case "GS":
                            wp2 = "Small Hail";
                            if (voiceAttackVariablesSave == true)
                            {
                                VoiceAttackPlugin.SetText("runwayState", "WET");
                                VoiceAttackPlugin.SetText("visibleMoisture", "true");
                            }
                            break;
                        case "HZ":
                            wp2 = "Haze";
                            break;
                        case "IC":
                            wp2 = "Ice Crystals";
                            if (voiceAttackVariablesSave == true)
                            {
                                VoiceAttackPlugin.SetText("runwayState", "WET");
                                VoiceAttackPlugin.SetText("visibleMoisture", "true");
                            }
                            break;
                        case "PL":
                            wp2 = "Ice Pellets";
                            if (voiceAttackVariablesSave == true)
                            {
                                VoiceAttackPlugin.SetText("runwayState", "WET");
                                VoiceAttackPlugin.SetText("visibleMoisture", "true");
                            }
                            break;
                        case "PO":
                            wp2 = "Dust/Sand Whirls";
                            break;
                        case "PY":
                            wp2 = "Spray";
                            break;
                        case "RA":
                            wp2 = "Rain";
                            if (voiceAttackVariablesSave == true)
                            {
                                VoiceAttackPlugin.SetText("runwayState", "WET");
                                VoiceAttackPlugin.SetText("visibleMoisture", "true");
                            }
                            break;
                        case "SA":
                            wp2 = "Sand";
                            break;
                        case "SG":
                            wp2 = "Snow Grains";
                            if (voiceAttackVariablesSave == true)
                            {
                                VoiceAttackPlugin.SetText("runwayState", "WET");
                                VoiceAttackPlugin.SetText("visibleMoisture", "true");
                            }
                            break;
                        case "SH":
                            wp2 = "Showers";
                            if (voiceAttackVariablesSave == true)
                            {
                                VoiceAttackPlugin.SetText("visibleMoisture", "true");
                            }
                            break;
                        case "SN":
                            wp2 = "Snow";
                            if (voiceAttackVariablesSave == true)
                            {
                                VoiceAttackPlugin.SetText("runwayState", "WET");
                                VoiceAttackPlugin.SetText("visibleMoisture", "true");
                            }
                            break;
                        case "SQ":
                            wp2 = "Squalls";
                            break;
                        case "SS":
                            wp2 = "Sandstorm";
                            break;
                        case "VA":
                            wp2 = "Volcanic Ash";
                            break;
                        case "UP":
                            wp2 = "Unknown Precipitation";
                            break;
                    }
                }

                if (descriptor != "")
                {
                    if (preDescriptor == 0)
                    {
                        weatherMessage = weatherMessage + " " + descriptor;
                    }
                    else if (preDescriptor == 1)
                    {
                        weatherMessage = descriptor + " " + weatherMessage;
                    }
                    else if (preDescriptor == -1)
                    {
                        weatherMessage = descriptor + " and " + weatherMessage;
                    }
                }

                if (intensity != "")
                {
                    weatherMessage = intensity + " " + weatherMessage;
                }

                if (weatherPhenomenaType.Length == 3)
                {
                    if (weatherPhenomenaType == "NSW")
                    {
                        weatherMessage = "No Significant Weather expected";
                    }

                }

                VoiceAttackPlugin.LogOutput("Report WP final message: " + weatherMessage, "grey");

                wpParts.Add(weatherMessage);

                if (wp2 != "")
                {
                    wpParts.Add(wp2);
                    VoiceAttackPlugin.LogOutput("Report WP second message: " + wp2, "grey");
                }
            }
        }

        // SKY CONDITIONS DETECTION
        public void SCcheck(string component)
        {
            VoiceAttackPlugin.SetText(reportPrefix + ".SCcount" + reportSufix, "-1");

            if (Regex.IsMatch(component, @"^(FEW|SCT|BKN|OVC|VV)\d{3}$|^(FEW|SCT|BKN|OVC|VV)\d{3}[A-Za-z]{2,3}$|^CAVOK$|^CLR$|^NSC$|^NCD$|^SKC$|^VV///$"))
            {
                VoiceAttackPlugin.LogOutput("Report SKY component identified: " + component, "grey");

                scExists = true;

                string scMessage = "0";

                string cloudState = "0";
                string cloudHeight = "0";
                string specialCase = "0";

                if (component.StartsWith("CLR"))
                {
                    cloudState = "Clear Skies"; // Clear
                    cloudHeight = "0";
                }
                else if (component.StartsWith("SKC"))
                {
                    cloudState = "Clear Skies"; // Clear
                    cloudHeight = "0";
                }
                else if (component.StartsWith("NSC"))
                {
                    cloudState = "No Significant Clouds"; // No significant clouds
                    cloudHeight = "0"; // No significant clouds 
                }
                else if (component.StartsWith("NCD"))
                {
                    cloudState = "No Significant Clouds"; // No significant clouds
                    cloudHeight = "0"; // No significant clouds 
                }
                else if (component.StartsWith("FEW"))
                {
                    cloudState = "Few Clouds"; // Few clouds
                    cloudHeight = component.Substring(3, 3); // Height in hundreds of feet
                }
                else if (component.StartsWith("SCT"))
                {
                    cloudState = "Scattered Clouds"; // Scattered clouds
                    cloudHeight = component.Substring(3, 3); // Height in hundreds of feet
                }
                else if (component.StartsWith("BKN"))
                {
                    cloudState = "Broken Clouds"; // Broken clouds
                    cloudHeight = component.Substring(3, 3); // Height in hundreds of feet
                }
                else if (component.StartsWith("OVC"))
                {
                    cloudState = "Overcast Clouds"; // Overcast clouds
                    cloudHeight = component.Substring(3, 3); // Height in hundreds of feet
                }
                else if (component.StartsWith("VV"))
                {
                    cloudState = "Vertical Visibility"; // Vertical Visibility
                    cloudHeight = component.Substring(2, 3); // Height in hundreds of feet
                }
                else if (component.StartsWith("CAVOK"))
                {
                    cloudState = "Cloud and visibility OK"; // CAVOK
                    cloudHeight = "0";
                }

                VoiceAttackPlugin.LogOutput("Report SKY component state: " + cloudState, "grey");
                VoiceAttackPlugin.LogOutput("Report SKY component height: " + cloudHeight, "grey");

                if (component.EndsWith("CB"))
                {
                    specialCase = "Cumulonimbus";
                    VoiceAttackPlugin.LogOutput("Report SKY component special: " + specialCase, "grey");
                }

                if (component.EndsWith("TCU"))
                {
                    specialCase = "Towering Cumulus";
                    VoiceAttackPlugin.LogOutput("Report SKY component special: " + specialCase, "grey");
                }

                if (cloudHeight != "0")
                {
                    if (cloudHeight == "///")
                    {
                        cloudHeight = "Undetermined Height";
                    }
                    else 
                    {
                        if (cloudHeight.StartsWith("0"))
                            cloudHeight = cloudHeight.Substring(1);
                        if (cloudHeight.StartsWith("0"))
                            cloudHeight = cloudHeight.Substring(1);
                        cloudHeight = cloudHeight + "00 feet AGL";
                        if (voiceAttackVariablesSave == true)
                        {
                            int intCloudHeight = int.Parse(cloudHeight.Substring(0, cloudHeight.Length - 9));
                            int intCloudHeightlatest = VoiceAttackPlugin.GetInt("lowestClouds");
                            if (intCloudHeight < intCloudHeightlatest)
                            {
                                VoiceAttackPlugin.SetInt("lowestClouds", intCloudHeight);
                            }
                        }
                    }                                    
                }

                if (specialCase != "0") 
                {
                    scMessage = cloudState + " at " + cloudHeight + " with " + specialCase + " above";
                }
                else
                {
                    if (cloudState == "Cloud and visibility OK" || cloudState == "No Significant Clouds" || cloudState == "Clear Skies")
                    {
                        scMessage = cloudState;
                    }
                    else
                    {
                        scMessage = cloudState + " at " + cloudHeight;
                    }                    
                }

                scParts.Add(scMessage);

            }
        }
    }
}
