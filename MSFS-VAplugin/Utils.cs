//=================================================================================================================
// PROJECT: MSFS VAPlugin
// PURPOSE: Contains misc. helper functions
// AUTHOR: William Riker
//================================================================================================================= 
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using Microsoft.FlightSimulator.SimConnect;
using static MSFS.VoiceAttackPlugin;

namespace MSFS
{

    public class Utils
    {
        public static string current = "NULL";

        public static string errvar = "NULL";

        public static string simbriefURL = "NULL";

        public static string pilotID = "NULL";

        public static string reqICAO = "NULL";

        public static string reqRunway = "NULL";

        public static string reqVORICAO = "NULL";

        public static string reqVORregion = "NULL";

        public static int simConnectMSGLoop = 1;

        public static string simvar = "NULL";

        public static string simvarindex = "NULL";

        public static string simvarunit = "NULL";

        public static SIMCONNECT_DATATYPE dataType = new SIMCONNECT_DATATYPE();

        public static string resultDataString = "NULL";

        public static double resultDataValue = 0;

        public static string resultDataType = "NULL";

        public static bool simVarSubscription = false;

        public static string calcString = "NULL";

        public static string facilityType = "NULL";

        public static string rmkRep = "";
        public static string tempoRep = "";
        public static string becmgRep = "";

        public static int TAFsection = 100;

        public static int wpQuantity = 0;

        public static bool tempoRepExists = false;
        public static bool becmgRepExists = false;
        public static bool rmkRepExists = false;
        public static bool stationIdentified = false;



        //SimBrief variables-----------------------------------------------------------------------

        public static int sbWeatherReportCharPos = 0;

        public static int sbWeatherReportCharPos2 = 0;

        public static string sbFlight = "NULL";

        public static string sbAirlineICAO = "NULL";

        public static string sbCallsign = "NULL";

        public static string sbCruiseProf = "NULL";

        public static string sbClimbProf = "NULL";

        public static string sbDescentProf = "NULL";

        public static int sbCostIndex = 0;

        public static int sbInitialAlt = 0;

        public static int sbAvgWindComp = 0;

        public static int sbAvgWindDir = 0;

        public static int sbAvgWindSpd = 0;

        public static double sbTopClimbOAT = 0;

        public static string sbRoute = "NULL";

        public static string sbOrigin = "NULL";

        public static int sbOriginElevation = 0;

        public static string sbOriginRwy = "NULL";

        public static int sbOriginTransAlt = 0;

        public static int sbOriginTransLevel = 0;

        public static int sbOriginAvailWeather = 0;

        public static string sbOriginWind = "NULL";

        public static string sbOriginMetar = "NULL";

        public static string sbOriginMetarTime = "NULL";

        public static string sbOriginTAF = "NULL";

        public static string sbOriginTAFTime = "NULL";

        public static string sbOriginPressure = "NULL";

        public static double sbOriginPressureINHG = 0;

        public static string sbOriginTemp = "NULL";

        public static string sbAltn = "NULL";

        public static int sbAltnElevation = 0;

        public static string sbAltnRwy = "NULL";

        public static int sbAltnTransAlt = 0;

        public static int sbAltnTransLevel = 0;

        public static int sbAltnAvailWeather = 0;

        public static string sbAltnMetar = "NULL";

        public static string sbAltnMetarTime = "NULL";

        public static string sbAltnTAF = "NULL";

        public static string sbAltnTAFTime = "NULL";

        public static string sbAltnWind = "NULL";

        public static string sbAltnPressure = "NULL";

        public static double sbAltnPressureINHG = 0;

        public static string sbAltnTemp = "NULL";

        public static string sbDestination = "NULL";

        public static int sbDestElevation = 0;

        public static string sbDestRwy = "NULL";

        public static int sbDestTransAlt = 0;

        public static int sbDestTransLevel = 0;

        public static int sbDestAvailWeather = 0;

        public static string sbDestMetar = "NULL";

        public static string sbDestMetarTime = "NULL";

        public static string sbDestTAF = "NULL";

        public static string sbDestTAFTime = "NULL";

        public static string sbDestWind = "NULL";

        public static string sbDestPressure = "NULL";

        public static double sbDestPressureINHG = 0;

        public static string sbDestTemp = "NULL";

        public static string sbUnits = "NULL";

        public static double sbFinRes = 0;

        public static double sbAltnFuel = 0;

        public static double sbFinresPAltn = 0;

        public static double sbFuel = 0;

        public static int sbPassenger = 0;

        public static int sbBags = 0;

        public static double sbWeightPax = 0;

        public static double sbWeightCargo = 0;

        public static double sbPayload = 0;

        public static double sbZFW = 0;

        public static double sbTOW = 0;


        //-----------------------------------------------------------------------------------------

        public static double scLatitude = 0;
        public static double scLongitude = 0;
        public static double scAltitude = 0;

        public static float scAltitude1 = 0;
        public static float scHeading1 = 0;
        public static float scLength1 = 0;
        public static float scSlope1 = 0;
        public static float scTrueSlope1 = 0;
        public static int scPrimNumb1 = 0;
        public static int scPrimDesign1 = 0;
        public static string scPrimVORICAO1 = "NULL";
        public static string scPrimVORregion1 = "NULL";
        public static int scSecNumb1 = 0;
        public static int scSecDesign1 = 0;
        public static string scSecVORICAO1 = "NULL";
        public static string scSecVORregion1 = "NULL";
        public static float scAltitude2 = 0;
        public static float scHeading2 = 0;
        public static float scLength2 = 0;
        public static float scSlope2 = 0;
        public static float scTrueSlope2 = 0;
        public static int scPrimNumb2 = 0;
        public static int scPrimDesign2 = 0;
        public static string scPrimVORICAO2 = "NULL";
        public static string scPrimVORregion2 = "NULL";
        public static int scSecNumb2 = 0;
        public static int scSecDesign2 = 0;
        public static string scSecVORICAO2 = "NULL";
        public static string scSecVORregion2 = "NULL";
        public static float scAltitude3 = 0;
        public static float scHeading3 = 0;
        public static float scLength3 = 0;
        public static float scSlope3 = 0;
        public static float scTrueSlope3 = 0;
        public static int scPrimNumb3 = 0;
        public static int scPrimDesign3 = 0;
        public static string scPrimVORICAO3 = "NULL";
        public static string scPrimVORregion3 = "NULL";
        public static int scSecNumb3 = 0;
        public static int scSecDesign3 = 0;
        public static string scSecVORICAO3 = "NULL";
        public static string scSecVORregion3 = "NULL";
        public static float scAltitude4 = 0;
        public static float scHeading4 = 0;
        public static float scLength4 = 0;
        public static float scSlope4 = 0;
        public static float scTrueSlope4 = 0;
        public static int scPrimNumb4 = 0;
        public static int scPrimDesign4 = 0;
        public static string scPrimVORICAO4 = "NULL";
        public static string scPrimVORregion4 = "NULL";
        public static int scSecNumb4 = 0;
        public static int scSecDesign4 = 0;
        public static string scSecVORICAO4 = "NULL";
        public static string scSecVORregion4 = "NULL";
        public static float scAltitude5 = 0;
        public static float scHeading5 = 0;
        public static float scLength5 = 0;
        public static float scSlope5 = 0;
        public static float scTrueSlope5 = 0;
        public static int scPrimNumb5 = 0;
        public static int scPrimDesign5 = 0;
        public static string scPrimVORICAO5 = "NULL";
        public static string scPrimVORregion5 = "NULL";
        public static int scSecNumb5 = 0;
        public static int scSecDesign5 = 0;
        public static string scSecVORICAO5 = "NULL";
        public static string scSecVORregion5 = "NULL";
        public static float scAltitude6 = 0;
        public static float scHeading6 = 0;
        public static float scLength6 = 0;
        public static float scSlope6 = 0;
        public static float scTrueSlope6 = 0;
        public static int scPrimNumb6 = 0;
        public static int scPrimDesign6 = 0;
        public static string scPrimVORICAO6 = "NULL";
        public static string scPrimVORregion6 = "NULL";
        public static int scSecNumb6 = 0;
        public static int scSecDesign6 = 0;
        public static string scSecVORICAO6 = "NULL";
        public static string scSecVORregion6 = "NULL";
        public static float scAltitude7 = 0;
        public static float scHeading7 = 0;
        public static float scLength7 = 0;
        public static float scSlope7 = 0;
        public static float scTrueSlope7 = 0;
        public static int scPrimNumb7 = 0;
        public static int scPrimDesign7 = 0;
        public static string scPrimVORICAO7 = "NULL";
        public static string scPrimVORregion7 = "NULL";
        public static int scSecNumb7 = 0;
        public static int scSecDesign7 = 0;
        public static string scSecVORICAO7 = "NULL";
        public static string scSecVORregion7 = "NULL";
        public static float scAltitude8 = 0;
        public static float scHeading8 = 0;
        public static float scLength8 = 0;
        public static float scSlope8 = 0;
        public static float scTrueSlope8 = 0;
        public static int scPrimNumb8 = 0;
        public static int scPrimDesign8 = 0;
        public static string scPrimVORICAO8 = "NULL";
        public static string scPrimVORregion8 = "NULL";
        public static int scSecNumb8 = 0;
        public static int scSecDesign8 = 0;
        public static string scSecVORICAO8 = "NULL";
        public static string scSecVORregion8 = "NULL";


        public static float scLocFreq = 0;
        public static float scLocHeading = 0;
        public static string scLocName = "NULL";

        //-----------------------------------------------------------------------------------------

        public static bool errcon = false;

        public static string webSocketMessage;

        public static string webSocketColor;


        public static void Calculator(string type)
        {
            //VoiceAttackPlugin.ForceLogOutput(type, "grey");

            if (type == "sin")
            {
                string value = VoiceAttackPlugin.GetText("MSFS.ValueSet");

                value = value.Replace(',', '.');

                double nValue = double.Parse(value);

                //VoiceAttackPlugin.ForceLogOutput(nValue.ToString(), "grey");

                double radians = nValue * (Math.PI / 180);

                //VoiceAttackPlugin.ForceLogOutput(radians.ToString(), "grey");

                double sineValue = Math.Sin(radians);

                //VoiceAttackPlugin.ForceLogOutput(sineValue.ToString(), "grey");

                decimal sineValueDecimal = (decimal)sineValue;

                //VoiceAttackPlugin.ForceLogOutput(sineValueDecimal.ToString(), "grey");

                VoiceAttackPlugin.SetDecimal("MSFS.ValueGet", sineValueDecimal);
            }

        }

        public static uint DecimalToScaledUInt(Decimal frequency)
        {
            return Decimal.ToUInt32(Math.Round(frequency * 1000)); // Scale by 1000 and round
        }

        public static void SetCallsign()
        {
            sbCallsign = ConvertToRadiophonic(sbFlight);

            sbCallsign = sbAirlineICAO + " " + sbCallsign;
        }

        public static string ConvertToRadiophonic(string toRadiophonic) 
        
        {
            toRadiophonic = toRadiophonic.ToUpper(); // Convert input to uppercase for consistent matching
            string converted = "";

            foreach (char c in toRadiophonic)
            {
                if (phoneticAlphabet.ContainsKey(c))
                    converted += phoneticAlphabet[c] + " ";
                else
                    converted += c + " "; // Use the original character if not found in the alphabet
            }

            return converted.Trim(); // Remove trailing space
        }


        public static uint DecimalToBCD16(uint num)
        {
            uint result = 0;
            uint shift = 0;

            while (num > 0)
            {
                uint digit = num % 10; // Extract the last decimal digit
                result |= (digit << (int)shift); // Encode the digit in BCD
                shift += 4; // Move to the next 4-bit slot
                num /= 10; // Remove the last digit
            }

            return result;
        }

        public static void SetKeyName(string KeyName)
        {
            current = KeyName;
        }
        public static double Deg2Rad(double deg)
        {
            return deg * Math.PI / 180;
        }

        public static double Rad2Deg(double rad)
        {
            return rad * 180 / Math.PI;
        }

        public static uint Bcd2Dec(uint num) 
        { 
            return HornerScheme(num, 0x10, 10); 
        }

        public static uint Dec2Bcd(uint num)
        {
            return HornerScheme(num, 10, 0x10);
        }

        private static uint HornerScheme(uint Num, uint Divider, uint Factor)
        {
            VoiceAttackPlugin.LogOutput("Horner scheme: " + Num + " Divider: " + Divider + " Factor: " + Factor, "grey");
            uint Remainder = 0, Quotient = 0, Result = 0;
            Remainder = Num % Divider; // Extract last digit

            VoiceAttackPlugin.LogOutput("Horner scheme remainder: " + Remainder, "grey");

            Quotient = Num / Divider; // Remove last digit

            VoiceAttackPlugin.LogOutput("Horner scheme quotient: " + Quotient, "grey");

            if (!(Quotient == 0 && Remainder == 0)) // Recurse if there's more digits
                Result += HornerScheme(Quotient, Divider, Factor) * Factor + Remainder;

            VoiceAttackPlugin.LogOutput("Horner scheme result: " + Result, "grey");

            return Result;
        }

        static Dictionary<char, string> phoneticAlphabet = new Dictionary<char, string>
    {
        { 'A', "Alpha" },
        { 'B', "Bravo" },
        { 'C', "Charlie" },
        { 'D', "Delta" },
        { 'E', "Echo" },
        { 'F', "Foxtrot" },
        { 'G', "Golf" },
        { 'H', "Hotel" },
        { 'I', "India" },
        { 'J', "Juliett" },
        { 'K', "Kilo" },
        { 'L', "Lima" },
        { 'M', "Mike" },
        { 'N', "November" },
        { 'O', "Oscar" },
        { 'P', "Papa" },
        { 'Q', "Quebec" },
        { 'R', "Romeo" },
        { 'S', "Sierra" },
        { 'T', "Tango" },
        { 'U', "Uniform" },
        { 'V', "Victor" },
        { 'W', "Whiskey" },
        { 'X', "X-ray" },
        { 'Y', "Yankee" },
        { 'Z', "Zulu" },
        { '9', "Niner" }

    };

    }
}
