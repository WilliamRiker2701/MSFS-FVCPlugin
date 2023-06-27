//=================================================================================================================
// PROJECT: MSFS VAPlugin
// PURPOSE: Contains misc. helper functions
// AUTHOR: William Riker
//================================================================================================================= 
using System;
using static MSFS.VoiceAttackPlugin;

namespace MSFS
{

    public class Utils
    {
        public static string current = "NULL";

        public static string errvar = "NULL";

        public static string simbriefURL = "NULL";

        public static string pilotID = "NULL";

        public static string xmlFile = "NULL";

        //SimBrief variables-----------------------------------------------------------------------

        public static string sbFlight = "NULL";

        public static int sbCostIndex = 0;

        public static string sbRoute = "NULL";

        public static string sbOrigin = "NULL";

        public static int sbOriginElevation = 0;

        public static string sbOriginRwy = "NULL";

        public static int sbOriginTransAlt = 0;

        public static int sbOriginTransLevel = 0;

        public static int sbOriginWindDir = 0;

        public static int sbOriginWindSpd = 0;

        public static string sbOriginMetar = "NULL";

        public static int sbOriginQNH = 0;

        public static double sbOriginBaro = 0;

        public static string sbAltn = "NULL";

        public static int sbAltnElevation = 0;

        public static string sbAltnRwy = "NULL";

        public static int sbAltnTransAlt = 0;

        public static int sbAltnTransLevel = 0;

        public static string sbAltnMetar = "NULL";

        public static int sbAltnWindDir = 0;

        public static int sbAltnWindSpd = 0;

        public static int sbAltnQNH = 0;

        public static double sbAltnBaro = 0;

        public static string sbDestination = "NULL";

        public static int sbDestElevation = 0;

        public static string sbDestRwy = "NULL";

        public static int sbDestTransAlt = 0;

        public static int sbDestTransLevel = 0;

        public static string sbDestMetar = "NULL";

        public static int sbDestWindDir = 0;

        public static int sbDestWindSpd = 0;

        public static int sbDestQNH = 0;

        public static double sbDestBaro = 0;

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

        public static bool errcon = false;



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
        
        public static uint Dec2Bcd(uint num) { 
            return HornerScheme(num, 10, 0x10); 
        } 
        
        static private uint HornerScheme(uint Num, uint Divider, uint Factor) 
        { 
            uint Remainder = 0, Quotient = 0, Result = 0; 
            Remainder = Num % Divider; 
            Quotient = Num / Divider; 
            
            if (!(Quotient == 0 && Remainder == 0)) 
                Result += HornerScheme(Quotient, Divider, Factor) * Factor + Remainder; 
            
            return Result; 
        }

    }
}
