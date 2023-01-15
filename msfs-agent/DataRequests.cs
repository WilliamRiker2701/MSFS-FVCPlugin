//=================================================================================================================
// PROJECT: MSFS Agent V2
// PURPOSE: This file defines the data structures used to request data from the sim.
// AUTHOR: James Clark and William Riker
// Licensed under the MS-PL license. See LICENSE.md file in the project root for full license information.
//================================================================================================================= 
using System.Runtime.InteropServices;

namespace MSFS
{
    //Defines the notification groups we want to use
    public enum NOTIFICATION_GROUPS
    {
        DEFAULT
    }

    public enum hSimconnect : int
    {
        group1
    }

    // Identifies the types of data requests we want to make
    public enum RequestTypes
    {
        PlaneState,
    }

    // Identifies the different data definitions we've defined
    public enum DataDefinitions
    {
        PlaneState,
    }


    // We need data structure for each data request
    // Note: each string in the data structure needs a MarshalAs statement above it
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi, Pack = 1)]
    public struct PlaneState
    {
        
        public int Airspeed_Indicated;
        public int Ambient_Temperature;
        public bool Apu_Switch;
        public bool Apu_Generator_Switch;
        public double Apu_Pct_Rpm;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
        public string Atc_Airline;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
        public string Atc_Flight_Number;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
        public string Atc_Id;
        public double Auto_Brake_Switch_CB;
        public bool Autopilot_Airspeed_Hold;
        public int Autopilot_Airspeed_Hold_Var;
        public bool Autopilot_Altitude_Lock;
        public int Autopilot_Altitude_Lock_Var;
        public bool Autopilot_Approach_Hold;
        public bool Autopilot_Attitude_Hold;
        public bool Autopilot_Available;
        public bool Autopilot_Backcourse_Hold;
        public bool Autopilot_Disengaged;
        public bool Autopilot_Flight_Director_Active;
        public bool Autopilot_Heading_Lock;
        public int Autopilot_Heading_Lock_Dir;
        public bool Autopilot_Master;
        public int Autopilot_Max_Bank_ID;
        public double Autopilot_Nav_Selected;
        public bool Autopilot_Nav1_Lock;
        public bool Autopilot_Throttle_Arm;
        public bool Autopilot_Vertical_Hold;
        public int Autopilot_Vertical_Hold_Var;
        public bool Autopilot_Yaw_Damper;
        public int Bleed_Air_Source_Control;
        public bool Brake_Parking_Indicator;
        public bool Circuit_Switch_On_17;
        public bool Circuit_Switch_On_18;
        public bool Circuit_Switch_On_19;
        public bool Circuit_Switch_On_20;
        public bool Circuit_Switch_On_21;
        public bool Circuit_Switch_On_22;
        public double Com1_Active_Frequency;
        public double Com1_Standby_Frequency;
        public double Com2_Active_Frequency;
        public double Com2_Standby_Frequency;
        public bool Electrical_Master_Battery;
        public bool External_Power_On;
        public bool Engine_Anti_Ice_1;
        public bool Engine_Anti_Ice_2;
        public double Engine_N1_RPM_1;
        public double Engine_N1_RPM_2;
        public double Engine_N2_RPM_1;
        public double Engine_N2_RPM_2;
        public int Engine_Type;
        public double Flaps_Handle_Index;
        public double Flaps_Handle_Percent;
        public bool Fuelsystem_Pump_Switch_1;
        public bool Fuelsystem_Pump_Switch_2;
        public bool Fuelsystem_Pump_Switch_3;
        public bool Fuelsystem_Pump_Switch_4;
        public bool Fuelsystem_Pump_Switch_5;
        public bool Fuelsystem_Pump_Switch_6;
        public bool Fuelsystem_Valve_Switch_1;
        public bool Fuelsystem_Valve_Switch_2;
        public bool Fuelsystem_Valve_Switch_3;
        public int Gear_Handle_Position;
        public bool General_Eng_Starter_1;
        public bool General_Eng_Starter_2;
        public int Ground_Velocity;
        public int Heading_Indicator;
        public bool Hydraulic_Switch;
        public bool Is_Gear_Retractable;
        public bool Light_Beacon;
        public bool Light_Cabin;
        public bool Light_Landing;
        public bool Light_Logo;
        public bool Light_Nav;
        public bool Light_Panel;
        public bool Light_Recognition;
        public bool Light_Strobe;
        public bool Light_Taxi;
        public bool Light_Wing;
        public int Local_Time;
        public bool Master_Ignition_Switch;
        public double Nav1_Active_Frequency;
        public double Nav1_Standby_Frequency;
        public double Nav2_Active_Frequency;
        public double Nav2_Standby_Frequency;
        public double Number_Of_Engines;
        public bool Panel_Anti_Ice_Switch;
        public bool Pitot_Heat;
        public int Plane_Alt_Above_Ground;
        public int Plane_Altitude;
        public double Plane_Latitude;
        public double Plane_Longitude;
        public bool Prop_Deice_Switch;
        public int Pushback_State;
        public bool Sim_On_Ground;
        public bool Spoiler_Available;
        public double Spoilers_Handle_Position;
        public bool Structural_Deice_Switch;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
        public string Title;
        public bool Transponder_Available;
        public int Transponder_Code;
        public int Vertical_Speed;
        public int Water_Rudder_Handle_Position;
        public bool Windshield_Deice_Switch;
        public double Alt_Variable;
        public string PMDG_Variable;
    }

    public enum SimVarList
    {
        //  0=Bool
        //  1=Knots
        //  2=Celsius
        //  3=Percent Over 100
        //  4=Number
        //  5=Feet
        //  6=Degrees
        //  7=Integer
        //  8=Feet/Minute
        //  9=Enum
        //  10=Frequency BCD16
        //  11=RPM
        //  12=Radians
        //  13=MHz
        //  14=Position
        //  15=BCO16
        //  16=Feet per second

        Airspeed_Indicated = 1,
        Ambient_Temperature = 2,
        Apu_Switch = 0,
        Apu_Generator_Switch = 0,
        Apu_Pct_Rpm = 3,
        Auto_Brake_Switch_CB = 4,
        Autopilot_Airspeed_Hold = 0,
        Autopilot_Airspeed_Hold_Var = 1,
        Autopilot_Altitude_Lock = 0,
        Autopilot_Altitude_Lock_Var = 5,
        Autopilot_Approach_Hold = 0,
        Autopilot_Attitude_Hold = 0,
        Autopilot_Available = 0,
        Autopilot_Backcourse_Hold = 0,
        Autopilot_Disengaged = 0,
        Autopilot_Flight_Director_Active = 0,
        Autopilot_Heading_Lock = 0,
        Autopilot_Heading_Lock_Dir = 6,
        Autopilot_Max_Bank_ID = 7,
        Autopilot_Nav_Selected = 4,
        Autopilot_Nav1_Lock = 0,
        Autopilot_Throttle_Arm = 0,
        Autopilot_Vertical_Hold = 0,
        Autopilot_Vertical_Hold_Var = 8,
        Autopilot_Yaw_Damper = 0,
        Bleed_Air_Source_Control = 9,
        Brake_Parking_Indicator = 0,
        Circuit_Switch_On_17 = 0,
        Circuit_Switch_On_18 = 0,
        Circuit_Switch_On_19 = 0,
        Circuit_Switch_On_20 = 0,
        Circuit_Switch_On_21 = 0,
        Circuit_Switch_On_22 = 0,
        Com_Active_Frequency_1 = 10,
        Com_Active_Frequency_2 = 10,
        Com_Standby_Frequency_1 = 10,
        Com_Standby_Frequency_2 = 10,
        Electrical_Master_Battery = 0,
        External_Power_On = 0,
        Eng_Anti_Ice_1 = 0,
        Eng_Anti_Ice_2 = 0,
        Eng_N1_RPM_1 = 11,
        Eng_N1_RPM_2 = 11,
        Eng_N2_RPM_1 = 11,
        Eng_N2_RPM_2 = 11,
        Engine_Type = 9,
        Flaps_Handle_Index = 4,
        Flaps_Handle_Percent = 3,
        Fuelsystem_Pump_Switch_1 = 0,
        Fuelsystem_Pump_Switch_2 = 0,
        Fuelsystem_Pump_Switch_3 = 0,
        Fuelsystem_Pump_Switch_4 = 0,
        Fuelsystem_Pump_Switch_5 = 0,
        Fuelsystem_Pump_Switch_6 = 0,
        Fuelsystem_Valve_Switch_1 = 0,
        Fuelsystem_Valve_Switch_2 = 0,
        Fuelsystem_Valve_Switch_3 = 0,
        Gear_Handle_Position = 0,
        General_Eng_Starter_1 = 0,
        General_Eng_Starter_2 = 0,
        Ground_Velocity = 1,
        Heading_Indicator = 12,
        Hydraulic_Switch = 0,
        Is_Gear_Retractable = 0,
        Light_Beacon = 0,
        Light_Cabin = 0,
        Light_Landing = 0,
        Light_Logo = 0,
        Light_Nav = 0,
        Light_Panel = 0,
        Light_Recognition = 0,
        Light_Strobe = 0,
        Light_Taxi = 0,
        Light_Wing = 0,
        Master_Ignition_Switch = 0,
        Nav_Active_Frequency_1 = 13,
        Nav_Active_Frequency_2 = 13,
        Nav_Standby_Frequency_1 = 13,
        Nav_Standby_Frequency_2 = 13,
        Number_Of_Engines = 4,
        Panel_Anti_Ice_Switch = 0,
        Pitot_Heat = 0,
        Plane_Alt_Above_Ground = 5,
        Plane_Altitude = 5,
        Plane_Latitude = 12,
        Plane_Longitude = 12,
        Prop_Deice_Switch = 0,
        Pushback_State = 9,
        Sim_On_Ground = 0,
        Spoiler_Available = 0,
        Spoilers_Handle_Position = 14,
        Structural_Deice_Switch = 0,
        Transponder_Available = 0,
        Transponder_Code = 15,
        Vertical_Speed = 16,
        Water_Rudder_Handle_Position = 3,
        Windshield_Deice_Switch = 0,

    }

    public class SimVarFunc
    {

        public string GiveSimVarUnit(int iD)
        {

            string Unit = "NULL";
            
            switch (iD)
            {

                case 0:

                    Unit = "Bool";

                    break;

                case 1:

                    Unit = "Knots";

                    break;

                case 2:

                    Unit = "Celsius";

                    break;

                case 3:

                    Unit = "Percent Over 100";

                    break;

                case 4:

                    Unit = "Number";

                    break;

                case 5:

                    Unit = "Feet";

                    break;

                case 6:

                    Unit = "Degrees";

                    break;

                case 7:

                    Unit = "Integer";

                    break;

                case 8:

                    Unit = "Feet/Minute";

                    break;

                case 9:

                    Unit = "Enum";

                    break;

                case 10:

                    Unit = "Frequency BCD16";

                    break;

                case 11:

                    Unit = "RPM";

                    break;

                case 12:

                    Unit = "Radians";

                    break;

                case 13:

                    Unit = "MHz";

                    break;

                case 14:

                    Unit = "Position";

                    break;

                case 15:

                    Unit = "BCO16";

                    break;
                    
                case 16:

                    Unit = "Feet per second";

                    break;


            }

            return Unit;
        }
        


    }



}
