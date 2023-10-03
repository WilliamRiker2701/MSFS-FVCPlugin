//=================================================================================================================
// PROJECT: MSFS VAPlugin
// PURPOSE: This file defines the data structures for SimVar variables.
// AUTHOR: William Riker
//================================================================================================================= 
using System.Runtime.InteropServices;
using Microsoft.FlightSimulator.SimConnect;

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
        SimVar,
    }
    
    // Identifies the different data definitions we've defined
    public enum DataDefinitions
    {
        SimVar,
    }
   
    public enum FacilityRequestTypes
    {
        FacilityData,
    }


    public enum SIMVAR_DEFINITION
    {
        Dummy = 0
    };

    public enum REQUEST_DEFINITON
    {
        Dummy = 0,
        Struct1
    };


    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    struct airport
    {
        public double latitude;
        public double longitude;
        public double altitude;
        public System.Int32 arrivals;
        public System.Int32 nRunways;

    };

    struct runway
    {
        public float heading;
        public float length;
        public float slope;
        public float trueslope;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 8)]
        public string primaryVORICAO;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 8)]
        public string primaryVORregion;
        public int primaryNumber;
        public int primaryDesignator;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 8)]
        public string secondaryVORICAO;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 8)]
        public string secondaryVORregion;
        public int secondaryNumber;
        public int secondaryDesignator;

    }

    struct vor
    {

        public double gsAltitude;
        public int hasGlideSlope;
        public uint frequency;
        public int type;
        public float localizer;
        public float glide_slope;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)]
        public string name;


    }


    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi, Pack = 1)]
    public struct SimVarString
    {

        public string simvar;

    }

    public struct SimVarBool
    {

        public bool simvar;

    }

    public struct SimVarDouble
    {

        public double simvar;

    }

    public struct SimVarFloat
    {

        public float simvar;

    }

    public struct SimVarInt
    {

        public int simvar;

    }

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
        //  17=Volts
        //  18=inHg
        //  19=Meters
        //  20=Kilograms
        //  21=Gallons
        //  22=Slugs per cubic feet

        AIRSPEED_INDICATED = 1,
        AMBIENT_DENSITY = 22,
        AMBIENT_IN_CLOUD = 0,
        AMBIENT_TEMPERATURE = 2,
        AMBIENT_PRECIP_RATE = 19,
        AMBIENT_PRECIP_STATE = 4,
        AMBIENT_PRESSURE = 18,
        AMBIENT_VISIBILITY = 19,
        AMBIENT_WIND_DIRECTION = 6,
        AMBIENT_WIND_VELOCITY = 1,
        APU_SWITCH = 0,
        APU_GENERATOR_SWITCH = 0,
        APU_PCT_RPM = 3,
        AUTO_BRAKE_SWITCH_CB = 4,
        AUTOPILOT_AIRSPEED_HOLD = 0,
        AUTOPILOT_AIRSPEED_HOLD_VAR = 1,
        AUTOPILOT_ALTITUDE_LOCK = 0,
        AUTOPILOT_ALTITUDE_LOCK_VAR = 5,
        AUTOPILOT_APPROACH_HOLD = 0,
        AUTOPILOT_ATTITUDE_HOLD = 0,
        AUTOPILOT_AVAILABLE = 0,
        AUTOPILOT_BACKCOURSE_HOLD = 0,
        AUTOPILOT_DISENGAGED = 0,
        AUTOPILOT_FLIGHT_DIRECTOR_ACTIVE = 0,
        AUTOPILOT_HEADING_LOCK = 0,
        AUTOPILOT_HEADING_LOCK_DIR = 6,
        AUTOPILOT_MAX_BANK_ID = 7,
        AUTOPILOT_NAV_SELECTED = 4,
        AUTOPILOT_NAV1_LOCK = 0,
        AUTOPILOT_TAKEOFF_POWER_ACTIVE = 0,
        AUTOPILOT_THROTTLE_ARM = 0,
        AUTOPILOT_VERTICAL_HOLD = 0,
        AUTOPILOT_VERTICAL_HOLD_VAR = 8,
        AUTOPILOT_YAW_DAMPER = 0,
        BLEED_AIR_APU = 0,
        BLEED_AIR_ENGINE = 0,
        BLEED_AIR_SOURCE_CONTROL = 9,
        BRAKE_PARKING_INDICATOR = 0,
        CABIN_SEATBELTS_ALERT_SWITCH = 0,
        CIRCUIT_POWER_SETTING = 3,
        CIRCUIT_SWITCH_ON = 0,
        CG_PERCENT = 3,
        COM_ACTIVE_FREQUENCY = 13,
        COM_STANDBY_FREQUENCY = 13,
        ELECTRICAL_MAIN_BUS_VOLTAGE = 17,
        ELECTRICAL_MASTER_BATTERY = 0,
        EXTERNAL_POWER_AVAILABLE = 0,
        EXTERNAL_POWER_ON = 0,
        EMPTY_WEIGHT = 20,
        ENG_ANTI_ICE = 0,
        ENG_N1_RPM = 11,
        ENG_N2_RPM = 11,
        ENGINE_TYPE = 9,
        FLAPS_HANDLE_INDEX = 4,
        FLAPS_HANDLE_PERCENT = 3,
        FUELSYSTEM_TANK_QUANTITY = 21,
        FUEL_TOTAL_QUANTITY_WEIGHT = 20,
        FUEL_WEIGHT_PER_GALLON = 20,
        FUELSYSTEM_PUMP_SWITCH = 0,
        FUELSYSTEM_VALVE_SWITCH = 0,
        GEAR_HANDLE_POSITION = 0,
        GENERAL_ENG_MASTER_ALTERNATOR = 0,
        GENERAL_ENG_STARTER = 0,
        GROUND_VELOCITY = 1,
        HEADING_INDICATOR = 12,
        HYDRAULIC_SWITCH = 0,
        INDICATED_ALTITUDE = 5,
        IS_GEAR_RETRACTABLE = 0,
        LIGHT_BEACON = 0,
        LIGHT_CABIN = 0,
        LIGHT_LANDING = 0,
        LIGHT_LOGO = 0,
        LIGHT_NAV = 0,
        LIGHT_PANEL = 0,
        LIGHT_RECOGNITION = 0,
        LIGHT_STROBE = 0,
        LIGHT_TAXI = 0,
        LIGHT_WING = 0,
        MASTER_IGNITION_SWITCH = 0,
        NAV_ACTIVE_FREQUENCY = 13,
        NAV_STANDBY_FREQUENCY = 13,
        NUMBER_OF_ENGINES = 4,
        PANEL_ANTI_ICE_SWITCH = 0,
        PITOT_HEAT = 0,
        PITOT_ICE_PCT = 3,
        PLANE_ALT_ABOVE_GROUND = 5,
        PLANE_ALTITUDE = 5,
        PLANE_LATITUDE = 12,
        PLANE_LONGITUDE = 12,
        PROP_DEICE_SWITCH = 0,
        PUSHBACK_STATE = 9,
        SEA_LEVEL_PRESSURE = 18,
        SIM_ON_GROUND = 0,
        SPOILERS_ARMED = 0,
        SPOILER_AVAILABLE = 0,
        SPOILERS_HANDLE_POSITION = 14,
        STRUCTURAL_DEICE_SWITCH = 0,
        STRUCTURAL_ICE_PCT = 3,
        SURFACE_INFO_VALID = 0,
        SURFACE_CONDITION = 9,
        SURFACE_TYPE = 9,
        TOTAL_AIR_TEMPERATURE = 2,
        TOTAL_WEIGHT = 20,
        TRANSPONDER_AVAILABLE = 0,
        TRANSPONDER_CODE = 13,
        TRANSPONDER_STATE = 9,
        TURB_ENG_IGNITION_SWITCH_EX1 = 9,
        VERTICAL_SPEED = 16,
        WATER_RUDDER_HANDLE_POSITION = 3,
        WINDSHIELD_DEICE_SWITCH = 0,

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
                    Utils.dataType = SIMCONNECT_DATATYPE.INT32;
                    Utils.resultDataType = "Bool";

                    break;

                case 1:

                    Unit = "Knots";
                    Utils.dataType = SIMCONNECT_DATATYPE.FLOAT32;
                    Utils.resultDataType = "Float";

                    break;

                case 2:

                    Unit = "Celsius";
                    Utils.dataType = SIMCONNECT_DATATYPE.FLOAT32;
                    Utils.resultDataType = "Float";

                    break;

                case 3:

                    Unit = "Percent Over 100";
                    Utils.dataType = SIMCONNECT_DATATYPE.FLOAT64;
                    Utils.resultDataType = "Double";

                    break;

                case 4:

                    Unit = "Number";
                    Utils.dataType = SIMCONNECT_DATATYPE.FLOAT64;
                    Utils.resultDataType = "Double";

                    break;

                case 5:

                    Unit = "Feet";
                    Utils.dataType = SIMCONNECT_DATATYPE.FLOAT32;
                    Utils.resultDataType = "Float";

                    break;

                case 6:

                    Unit = "Degrees";
                    Utils.dataType = SIMCONNECT_DATATYPE.FLOAT32;
                    Utils.resultDataType = "Float";

                    break;

                case 7:

                    Unit = "Integer";
                    Utils.dataType = SIMCONNECT_DATATYPE.INT32;
                    Utils.resultDataType = "Int";

                    break;

                case 8:

                    Unit = "Feet/Minute";
                    Utils.dataType = SIMCONNECT_DATATYPE.FLOAT32;
                    Utils.resultDataType = "Float";

                    break;

                case 9:

                    Unit = "Enum";
                    Utils.dataType = SIMCONNECT_DATATYPE.INT32;
                    Utils.resultDataType = "Int";

                    break;

                case 10:

                    Unit = "Frequency BCD16";
                    Utils.dataType = SIMCONNECT_DATATYPE.FLOAT64;
                    Utils.resultDataType = "Double";

                    break;

                case 11:

                    Unit = "RPM";
                    Utils.dataType = SIMCONNECT_DATATYPE.FLOAT32;
                    Utils.resultDataType = "Float";

                    break;

                case 12:

                    Unit = "Radians";
                    Utils.dataType = SIMCONNECT_DATATYPE.FLOAT32;
                    Utils.resultDataType = "Float";

                    break;

                case 13:

                    Unit = "MHz";
                    Utils.dataType = SIMCONNECT_DATATYPE.FLOAT64;
                    Utils.resultDataType = "Double";

                    break;

                case 14:

                    Unit = "Position";
                    Utils.dataType = SIMCONNECT_DATATYPE.FLOAT64;
                    Utils.resultDataType = "Double";

                    break;

                case 15:

                    Unit = "BCO16";
                    Utils.dataType = SIMCONNECT_DATATYPE.FLOAT32;
                    Utils.resultDataType = "Float";

                    break;
                    
                case 16:

                    Unit = "Feet per second";
                    Utils.dataType = SIMCONNECT_DATATYPE.FLOAT32;
                    Utils.resultDataType = "Float";

                    break;

                case 17:

                    Unit = "Volts";
                    Utils.dataType = SIMCONNECT_DATATYPE.FLOAT32;
                    Utils.resultDataType = "Float";

                    break;

                case 18:

                    Unit = "inHg";
                    Utils.dataType = SIMCONNECT_DATATYPE.FLOAT32;
                    Utils.resultDataType = "Float";

                    break;

                case 19:

                    Unit = "Meters";
                    Utils.dataType = SIMCONNECT_DATATYPE.FLOAT32;
                    Utils.resultDataType = "Float";

                    break;

                case 20:

                    Unit = "Kilograms";
                    Utils.dataType = SIMCONNECT_DATATYPE.FLOAT32;
                    Utils.resultDataType = "Float";

                    break;

                case 21:

                    Unit = "Gallons";
                    Utils.dataType = SIMCONNECT_DATATYPE.FLOAT32;
                    Utils.resultDataType = "Float";

                    break;

                case 22:

                    Unit = "Slugs per cubic feet";
                    Utils.dataType = SIMCONNECT_DATATYPE.FLOAT32;
                    Utils.resultDataType = "Float";

                    break;


            }

            return Unit;
        }
        


    }



}
