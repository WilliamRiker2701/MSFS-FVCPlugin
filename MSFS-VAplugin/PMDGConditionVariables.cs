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
            APU_EGTNeedle = 0x64E8,
            ELEC_DCMeterSelector = 0x64A5,
            APU_Selector = 0x659B,
            ELEC_APUGenSw = 0x64B3,
            ELEC_annunGEN_BUS_OFF = 0x64B9,
            ELEC_annunAPU_GEN_OFF_BUS = 0x64BB,
            ELEC_BusPowered = 0x64D6,
            FUEL_CrossFeedSw = 0x6478,
            ICE_WindowHeatTestSw = 0x6520,
            ADF_StandbyFrequency = 0x6470,
            FUEL_PumpFwdSw = 0x6479,
            ENG_EECSwitch = 0x6444,
            ENG_StartValve = 0x644C




        }
        public enum PMDGVarTypes
        {
            APU_EGTNeedle = 41,
            ELEC_DCMeterSelector = 11,
            APU_Selector = 11,
            ELEC_APUGenSw = 12,
            ELEC_annunGEN_BUS_OFF = 12,
            ELEC_BusPowered = 116,
            ELEC_annunAPU_GEN_OFF_BUS = 11,
            FUEL_CrossFeedSw = 11,
            ICE_WindowHeatTestSw = 22,
            ADF_StandbyFrequency = 21,
            FUEL_PumpFwdSw = 12,
            ENG_EECSwitch = 12,
            ENG_StartValve = 12


        }


    }
    
}
