//=================================================================================================================
// PROJECT: MSFS VAPlugin
// PURPOSE: This file defines the data structures for PMDG variables.
// AUTHOR: William Riker
//================================================================================================================= 

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MSFS
{

    public class PMDGConditionVariables
    {
        public enum PMDGVarAddress
        {
            AIR_FltAltWindow = 0x656C,
            AIR_LandAltWindow = 0x6572,
            APU_EGTNeedle = 0x64E8,
            ELEC_DCMeterSelector = 0x64A5,
            APU_Selector = 0x659B,
            ELEC_APUGenSw = 0x64B3,
            ELEC_annunGEN_BUS_OFF = 0x64B9,
            ELEC_annunAPU_GEN_OFF_BUS = 0x64BB,
            ELEC_annunGRD_POWER_AVAILABLE = 0x64AE,
            ELEC_GrdPwrSw = 0x64AF,
            ELEC_BusPowered = 0x64D6,
            FUEL_CrossFeedSw = 0x6478,
            ICE_WindowHeatTestSw = 0x6520,
            ADF_StandbyFrequency = 0x6470,
            FUEL_PumpFwdSw = 0x6479,
            ENG_EECSwitch = 0x6444,
            ENG_StartValve = 0x644C,
            MCP_Heading = 0x652C,
            MCP_Altitude = 0x652E,
            MCP_VertSpeed = 0x6530,
            MCP_IASMach = 0x6524,
            LTS_PedPanelKnob = 0x65D0,
            FUEL_annunLOWPRESS_Fwd = 0x648D,
            FUEL_annunLOWPRESS_Aft = 0x648F,
            FUEL_annunLOWPRESS_Ctr = 0x6491,




        }
        public enum PMDGVarTypes
        {
            AIR_FltAltWindow = 56,
            AIR_LandAltWindow = 56,
            APU_EGTNeedle = 41,
            ELEC_DCMeterSelector = 11,
            APU_Selector = 11,
            ELEC_APUGenSw = 12,
            ELEC_annunGEN_BUS_OFF = 12,
            ELEC_BusPowered = 116,
            ELEC_annunAPU_GEN_OFF_BUS = 11,
            ELEC_annunGRD_POWER_AVAILABLE = 11,
            ELEC_GrdPwrSw = 11,
            FUEL_CrossFeedSw = 11,
            ICE_WindowHeatTestSw = 22,
            ADF_StandbyFrequency = 21,
            FUEL_PumpFwdSw = 12,
            ENG_EECSwitch = 12,
            ENG_StartValve = 12,
            MCP_Heading = 611,
            MCP_Altitude = 611,
            MCP_VertSpeed = 20,
            MCP_IASMach = 42,
            LTS_PedPanelKnob = 20,
            FUEL_annunLOWPRESS_Fwd = 12,
            FUEL_annunLOWPRESS_Aft = 12,
            FUEL_annunLOWPRESS_Ctr = 12,


        }


    }
    
}
