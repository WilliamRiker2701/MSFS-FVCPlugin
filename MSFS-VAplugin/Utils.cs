//=================================================================================================================
// PROJECT: MSFS VAPlugin
// PURPOSE: Contains misc. helper functions
// AUTHOR: William Riker
//================================================================================================================= 
using System;

namespace MSFS
{

    public class Utils
    {
        public static string current = "NULL";

        public static string errvar = "NULL";

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
