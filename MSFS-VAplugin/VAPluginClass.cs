//=================================================================================================================
// PROJECT: MSFS VA Plugin
// PURPOSE: This class adds Microsoft Flight Simulator 2020 full support for Voice Attack 
// AUTHOR: William Riker
//================================================================================================================= 

using System;
using System.Threading;
using System.Diagnostics;
using WASimCommander.CLI.Enums;
using WASimCommander.CLI.Structs;
using WASimCommander.Client;
using WASimCommander.Enums;
using WASimCommander.CLI.Client;
using WASimCommander.CLI;
using System.Runtime.Remoting.Contexts;
using static System.Net.Mime.MediaTypeNames;
using System.Collections.Concurrent;
using FSUIPC;
using System.Configuration;

namespace MSFS
{
    public class VoiceAttackPlugin
    {
        const string VARIABLE_NAMESPACE = "MSFS";
        const string LOG_PREFIX = "MSFS: ";
        const string LOG_NORMAL = "purple";
        const string LOG_ERROR = "red";
        const string LOG_INFO = "grey";
        private static dynamic VA;



        /// <summary>
        /// Name of the plug-in as it should be shown in the UX
        /// </summary>
        public static string VA_DisplayName()
        {
            return "MSFS-FVCplugin - v1.1";
        }

        /// <summary>
        /// Extra information to display about the plug-in
        /// </summary>
        /// <returns></returns>
        public static string VA_DisplayInfo()
        {
            return "Enables the ability to trigger any in game actions/events vía SimConnect and WASM.";
        }

        /// <summary>
        /// Uniquely identifies the plugin
        /// </summary>
        public static Guid VA_Id()
        {
            return new Guid("{f7cc5aee-1590-4627-9407-d1aeb8d429d4}");
        }

        /// <summary>
        /// Used to stop any long running processes inside the plugin
        /// </summary>
        public static void VA_StopCommand()
        {
            // plugin has no long running processes
        }

        /// <summary>
        /// Runs when Voice Attack loads and processes plugins (runs once when the app launches)
        /// </summary>
        public static void VA_Init1(dynamic vaProxy)
        {

            // uncomment this line to force the debugger to attach at the very start of the class being created
            //System.Diagnostics.Debugger.Launch();
            VA = vaProxy;


        }

        /// <summary>
        /// Handles clean up before Voice Attack closes
        /// </summary>
        public static void VA_Exit1(dynamic vaProxy)
        {
            // no clean up needed
        }

        /// <summary>
        /// Main function used to process commands from Voice Attack
        /// </summary>
        public static void VA_Invoke1(dynamic vaProxy)
        {
            string context;
            string targetVar;
            string eventData;
            string eventData2;
            Commander msfsCommander;
            string varKind = "K";
            string varAction = "Set";
            
            double startValue;
            double curValue;
            double jumpValue;
            double varResult2;
            int slpValue;
            int i;
            int multInt;
            string multSt;
            string thrSleep;
            int aFrom;
            int aTo;
            bool sbFres;
            


            if (String.IsNullOrEmpty(vaProxy.Context))
                return;

            context = vaProxy.Context;

            if (context.Substring(0, 2) == "L:")
            {

                varKind = "L";

                if (context.Substring(context.Length - 2) == "-G")
                {

                    varAction = "Get";

                }

                if (context.Substring(context.Length - 2) == "-T")
                {

                    varAction = "Target";

                }
            }

            else if (context.Substring(0, 2) == "L-")
            {

                varKind = "L volume";

                if (context.Substring(context.Length - 2) == "-A")
                {

                    varAction = "Add";

                }

                else if (context.Substring(context.Length - 2) == "-S")
                {

                    varAction = "Substract";

                }

                else if (context.Substring(context.Length - 2) == "-R")
                {

                    varAction = "Repetition";

                }

            }

            else if (context.Substring(0, 2) == "P:")
            {

                varKind = "PMDG";

                if (context.Substring(context.Length - 2) == "-G")
                {

                    varAction = "Get";

                }

            }
            
            else if (context.Substring(0, 2) == "A:")
            {

                varKind = "A";
                                
            }

            else if (context.Substring(0, 2) == "K:")
            {

                varKind = "K";

            }

            else if (context.Substring(0, 3) == "CC:")
            {

                varKind = "CalcCode";

            }

            else if (context.Substring(0, 4) == "COM:")
            {

                varKind = "COM Key";

            }

            else if (context == "SimBriefFetch")
            {

                varKind = "SBF";

            }

            switch (varKind)
            {
                case "L":

                    switch (varAction)
                    {
                        case "Set":

                            double varResult;

                            double dData;

                            MSFS.Utils.errvar = context;

                            msfsCommander = ConnectToWASM(vaProxy);

                            while (MSFS.Utils.errcon == true)
                            {

                                vaProxy.WriteToLog(LOG_PREFIX + "Retrying connection", LOG_INFO);

                                msfsCommander = ConnectToWASM(vaProxy);

                            }
                                                                                                           
                            context = context.Remove(0, 2);

                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Processing Local Variable new value set: " + context, LOG_INFO);

                            eventData = vaProxy.GetText(VARIABLE_NAMESPACE + ".ValueSet");

                            dData = Convert.ToDouble(eventData);

                            msfsCommander.SetLVarFSUIPC(context, eventData);

                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "SetLVarFSUIPC Executed", LOG_INFO);

                            Thread.Sleep(200);

                            varResult = msfsCommander.GetLVarFSUIPC(context);

                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "New Value: " + varResult + "| Variable: " + context, LOG_NORMAL);

                            while (dData != varResult)
                            {
                                msfsCommander.SetLVarFSUIPC(context, eventData);

                                Thread.Sleep(200);

                                varResult = msfsCommander.GetLVarFSUIPC(context);

                                if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Set retry, New value: " + varResult, LOG_NORMAL);

                            }




                            if (MonitoringMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "New value: " + varResult + "| Variable: " + context, LOG_NORMAL);


                            msfsCommander.WASMDisconnect();                           
                           

                            break;

                        case "Target":

                            
                            MSFS.Utils.errvar = context;

                            msfsCommander = ConnectToWASM(vaProxy);

                            while (MSFS.Utils.errcon == true)
                            {

                                vaProxy.WriteToLog(LOG_PREFIX + "Retrying connection", LOG_INFO);

                                msfsCommander = ConnectToWASM(vaProxy);

                            }

                            context = context.Remove(0, 2);

                            context = context.Remove(context.Length - 2);

                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Processing Local Variable new value set: " + context, LOG_INFO);

                            targetVar = vaProxy.GetText(VARIABLE_NAMESPACE + ".TargetVar");

                            eventData2 = vaProxy.GetText(VARIABLE_NAMESPACE + ".ValueTarget");

                            eventData = vaProxy.GetText(VARIABLE_NAMESPACE + ".ValueSet");

                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Target variable: " + targetVar, LOG_INFO);

                            dData = Convert.ToDouble(eventData2);

                            msfsCommander.SetLVarFSUIPC(context, eventData);

                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "SetLVarFSUIPC Executed", LOG_INFO);

                            Thread.Sleep(200);

                            varResult = msfsCommander.GetLVarFSUIPC(context);
                            varResult2 = msfsCommander.GetLVarFSUIPC("L:"+targetVar);

                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "New Value: " + varResult + "| Variable: " + context, LOG_NORMAL);
                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Target Value: " + varResult2 + "| Target: " + targetVar, LOG_NORMAL);

                            while (dData != varResult2)
                            {
                                msfsCommander.SetLVarFSUIPC(context, eventData);

                                Thread.Sleep(200);

                                varResult2 = msfsCommander.GetLVarFSUIPC("L:" + targetVar);

                                if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Set retry, New target value: " + varResult2, LOG_NORMAL);

                            }




                            if (MonitoringMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "New value: " + varResult + "| Variable: " + context, LOG_NORMAL);


                            msfsCommander.WASMDisconnect();


                            break;

                        case "Get":

                            

                            string lVarSt = "NULL";

                            MSFS.Utils.errvar = context;

                            msfsCommander = ConnectToWASM(vaProxy);

                            while (MSFS.Utils.errcon == true)
                            {

                                vaProxy.WriteToLog(LOG_PREFIX + "Retrying connection", LOG_INFO);

                                msfsCommander = ConnectToWASM(vaProxy);

                            }

                            context = context.Remove(0, 2);

                            context = context.Remove(context.Length - 2);

                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Processing Local Variable data request: " + context, LOG_INFO);
                                                     
                            varResult = msfsCommander.GetLVarFSUIPC(context);

                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "varResult: " + varResult, LOG_INFO);

                            //if (msfsCommander.readsSep == true)
                            //{

                            lVarSt = varResult.ToString();

                                lVarSt = string.Join<char>(" ", lVarSt).Replace("- ", "-");

                                lVarSt = lVarSt.Replace(".", "point");

                                vaProxy.SetText(VARIABLE_NAMESPACE + ".ValueGetSt", lVarSt);

                                if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Local Variable request returns string: " + lVarSt, LOG_INFO);

                            //}

                            vaProxy.SetDecimal(VARIABLE_NAMESPACE + ".ValueGet", (decimal?)(double?)varResult);

                            string calcCode = "(L:" + context + ")";

                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Local Variable request returns value: " + varResult, LOG_INFO);

                            if (MonitoringMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Present value: " + varResult + "| Variable: " + context, LOG_NORMAL);
                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Present value: " + varResult + "| Variable: " + context, LOG_NORMAL);

                            msfsCommander.WASMDisconnect();

                            break;

                    }

                    break;

                case "L volume":


                    switch (varAction)
                    {
                        case "Add":

                            MSFS.Utils.errvar = context;

                            msfsCommander = ConnectToWASM(vaProxy);

                            while (MSFS.Utils.errcon == true)
                            {

                                vaProxy.WriteToLog(LOG_PREFIX + "Retrying connection", LOG_INFO);

                                msfsCommander = ConnectToWASM(vaProxy);

                            }

                            aFrom = context.IndexOf("L-") + "L-".Length;
                            aTo = context.IndexOf(":");

                            multSt = context.Substring(aFrom, aTo - aFrom);

                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Addition loop times: " + multSt, LOG_INFO);

                            multInt = int.Parse(multSt);

                            i = 0;

                            context = context.Remove(0, aTo+1);

                            context = context.Remove(context.Length - 2);

                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Processing Local Variable new value set: " + context, LOG_INFO);

                            eventData = vaProxy.GetText(VARIABLE_NAMESPACE + ".ValueSet");

                            thrSleep = vaProxy.GetText(VARIABLE_NAMESPACE + ".ThreadSleep");

                            jumpValue = double.Parse(eventData);

                            slpValue = int.Parse(thrSleep);

                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Set thread sleep (ms) at: " + slpValue, LOG_INFO);

                            startValue = msfsCommander.TriggerWASM(context);

                            curValue = startValue + jumpValue;

                            eventData = curValue.ToString();

                            while (i < multInt)
                            {


                                msfsCommander.TriggerWASM(context, eventData);

                                curValue = curValue + jumpValue;

                                eventData = curValue.ToString();

                                i = i + 1;

                                Thread.Sleep(slpValue);

                            }

                            if (MonitoringMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Addition value: " + jumpValue + "| Addition loop: " + multInt + " | Variable: " + context, LOG_NORMAL);
                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Addition value: " + jumpValue + "| Addition loop: " + multInt + " | Variable: " + context, LOG_NORMAL);

                            msfsCommander.WASMDisconnect();

                            break;

                        case "Substract":

                            MSFS.Utils.errvar = context;

                            msfsCommander = ConnectToWASM(vaProxy);

                            while (MSFS.Utils.errcon == true)
                            {

                                vaProxy.WriteToLog(LOG_PREFIX + "Retrying connection", LOG_INFO);

                                msfsCommander = ConnectToWASM(vaProxy);

                            }

                            aFrom = context.IndexOf("L-") + "L-".Length;
                            aTo = context.IndexOf(":");

                            multSt = context.Substring(aFrom, aTo - aFrom);

                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Substraction loop times: " + multSt, LOG_INFO);

                            multInt = int.Parse(multSt);

                            i = 0;

                            context = context.Remove(0, aTo+1);

                            context = context.Remove(context.Length - 2);

                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Processing Local Variable new value set: " + context, LOG_INFO);

                            eventData = vaProxy.GetText(VARIABLE_NAMESPACE + ".ValueSet");

                            thrSleep = vaProxy.GetText(VARIABLE_NAMESPACE + ".ThreadSleep");

                            jumpValue = double.Parse(eventData);

                            slpValue = int.Parse(thrSleep);

                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Set thread sleep (ms) at: " + slpValue, LOG_INFO);

                            if (MonitoringMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Thread sleep set: " + slpValue, LOG_NORMAL);

                            startValue = msfsCommander.TriggerWASM(context);

                            curValue = startValue - jumpValue;

                            eventData = curValue.ToString();


                            while (i < multInt)
                            {


                                msfsCommander.TriggerWASM(context, eventData);

                                curValue = curValue - jumpValue;

                                eventData = curValue.ToString();

                                i = i + 1;

                                Thread.Sleep(slpValue);

                            }

                            if (MonitoringMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Substraction value: " + jumpValue + "| Substraction loop: " + multInt + " | Variable: " + context, LOG_NORMAL);
                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Substraction value: " + jumpValue + "| Substraction loop: " + multInt + " | Variable: " + context, LOG_NORMAL);

                            msfsCommander.WASMDisconnect();

                            break;

                        case "Repetition":

                            MSFS.Utils.errvar = context;

                            msfsCommander = ConnectToWASM(vaProxy);

                            while (MSFS.Utils.errcon == true)
                            {

                                vaProxy.WriteToLog(LOG_PREFIX + "Retrying connection", LOG_INFO);

                                msfsCommander = ConnectToWASM(vaProxy);

                            }

                            aFrom = context.IndexOf("L-") + "L-".Length;
                            aTo = context.IndexOf(":");

                            multSt = context.Substring(aFrom, aTo - aFrom);

                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Repetition loop times: " + multSt, LOG_INFO);

                            multInt = int.Parse(multSt);

                            i = 0;

                            context = context.Remove(0, aTo+1);

                            context = context.Remove(context.Length - 2);

                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Processing Local Variable new value set: " + context, LOG_INFO);

                            eventData = vaProxy.GetText(VARIABLE_NAMESPACE + ".ValueSet");

                            thrSleep = vaProxy.GetText(VARIABLE_NAMESPACE + ".ThreadSleep");

                            while (i < multInt)
                            {

                                msfsCommander.TriggerWASM(context, eventData);

                                i = i + 1;

                            }

                            if (MonitoringMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Repetition value: " + eventData + "| Repetition loop: " + multInt + " | Variable: " + context, LOG_NORMAL);
                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Repetition value: " + eventData + "| Repetition loop: " + multInt + " | Variable: " + context, LOG_NORMAL);

                            msfsCommander.WASMDisconnect();

                            break;
                    } 
                    break;


                case "PMDG":

                    switch (varAction)
                    {
                        case "Set":

                            msfsCommander = ConnectToSim(vaProxy);

                            if (msfsCommander == null) return;

                            context = context.Remove(0, 2);

                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Processing PMDG Variable new value set: " + context, LOG_INFO);

                            eventData = vaProxy.GetText(VARIABLE_NAMESPACE + ".ValueSet");

                            msfsCommander.TriggerPMDG(context, eventData);

                            if (MonitoringMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "New value: " + eventData + "| Variable: " + context, LOG_NORMAL);
                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "New value: " + eventData + "| Variable: " + context, LOG_NORMAL);

                            msfsCommander.Disconnect();

                            break;

                        case "Get":

                            string varResult;

                            msfsCommander = ConnectToSim(vaProxy);

                            if (msfsCommander == null) return;

                            context = context.Remove(0, 2);

                            context = context.Remove(context.Length - 2);

                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Processing PMDG Variable data request: " + context, LOG_INFO);

                            varResult = msfsCommander.TriggerReqPMDG(context);

                            vaProxy.SetText(VARIABLE_NAMESPACE + ".ValueGet", varResult);

                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "FSUIPC data request returns: " + varResult, LOG_INFO);

                            if (MonitoringMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Present value: " + varResult + "| Variable: " + context, LOG_NORMAL);
                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Present value: " + varResult + "| Variable: " + context, LOG_NORMAL);

                            msfsCommander.Disconnect();

                            break;

                    }

                    break;

                case "K":

                    FsControlList reqKey;

                    context = context.Remove(0, 2);

                    Utils.SetKeyName(context);

                    msfsCommander = ConnectToSim(vaProxy);

                    if (msfsCommander == null) return;

                    if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Processing Key Event: " + context, LOG_INFO);

                    eventData = vaProxy.GetText(VARIABLE_NAMESPACE + ".ValueSet");

                    reqKey = (FsControlList)Enum.Parse(typeof(FsControlList), context, false);

                    msfsCommander.TriggerKeySimconnect(reqKey, eventData);

                    if (MonitoringMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "New value: " + eventData + "| Variable: " + context, LOG_NORMAL);
                    if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "New value: " + eventData + "| Variable: " + context, LOG_NORMAL);

                    msfsCommander.Disconnect();

                    break;

                case "COM Key":
                                        
                    MSFS.Utils.errvar = context;

                    msfsCommander = ConnectToWASM(vaProxy);

                    while (MSFS.Utils.errcon == true)
                    {

                        vaProxy.WriteToLog(LOG_PREFIX + "Retrying connection", LOG_INFO);

                        msfsCommander = ConnectToWASM(vaProxy);

                    }

                    context = context.Remove(0, 4);

                    if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Processing Key Event: " + context, LOG_INFO);

                    eventData = vaProxy.GetText(VARIABLE_NAMESPACE + ".ValueSet");

                    msfsCommander.TriggerComKey(context, eventData);

                    if (MonitoringMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "New value: " + eventData + "| Variable: " + context, LOG_NORMAL);
                    if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "New value: " + eventData + "| Variable: " + context, LOG_NORMAL);

                    msfsCommander.WASMDisconnect();

                    break;

                case "A":

                    double varRes;

                    string varResSt = "NULL";

                    MSFS.Utils.errvar = context;

                    msfsCommander = ConnectToWASM(vaProxy);

                    while (MSFS.Utils.errcon == true)
                    {

                        vaProxy.WriteToLog(LOG_PREFIX + "Retrying connection", LOG_INFO);

                        msfsCommander = ConnectToWASM(vaProxy);

                    }

                    context = context.Remove(0, 2);

                    if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Processing SimVar: " + context, LOG_INFO);

                    varRes = msfsCommander.TriggerReqSimVar(context);
                    
                    if (msfsCommander.readsSep == true)
                    {

                        varResSt = varRes.ToString();

                        varResSt = string.Join<char>(" ", varResSt).Replace("- ", "-");

                        varResSt = varResSt.Replace(".", "point");

                        vaProxy.SetText(VARIABLE_NAMESPACE + ".ValueGetSt", varResSt);

                        if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "SimVar request returns string: " + varResSt, LOG_NORMAL);

                    }

                    vaProxy.SetDecimal(VARIABLE_NAMESPACE + ".ValueGet", (decimal?)(double?)varRes);

                    if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "SimVar request returns value: " + varRes, LOG_NORMAL);

                    if (MonitoringMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Present value: " + varRes + "| Variable: " + context, LOG_NORMAL);
                    if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Present value: " + varRes + "| Variable: " + context, LOG_NORMAL);

                    msfsCommander.WASMDisconnect();

                    break;

                case "CalcCode":


                    MSFS.Utils.errvar = context;

                    msfsCommander = ConnectToWASM(vaProxy);

                    while (MSFS.Utils.errcon == true)
                    {

                        vaProxy.WriteToLog(LOG_PREFIX + "Retrying connection", LOG_INFO);

                        msfsCommander = ConnectToWASM(vaProxy);

                    }


                    context = context.Remove(0, 3);

                    if (MonitoringMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Sending calculator code: " + context, LOG_INFO);
                    if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Sending calculator code: " + context, LOG_INFO);

                    varRes = msfsCommander.TriggerCalcCode(context);

                    vaProxy.SetDecimal(VARIABLE_NAMESPACE + ".ValueGet", (decimal?)(double?)varRes);

                    if (MonitoringMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Calculator code response: " + varRes, LOG_NORMAL);
                    if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Calculator code response: " + varRes, LOG_NORMAL);

                    msfsCommander.WASMDisconnect();
                    

                    break;

                case "SBF":


                    msfsCommander = ConnectToWASM(vaProxy);

                    if (msfsCommander == null) return;

                    if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Processing SimBrief request...", LOG_INFO);

                    //string simbriefURL = ConfigurationManager.AppSettings["SBURL"].ToString();
                    //string pilotID = ConfigurationManager.AppSettings["pilotID"].ToString();
                    //string xmlFile = ConfigurationManager.AppSettings["xmlFile"].ToString();

                    new ApplicationConfigurationReader().Read();
                    var simbriefURL = ConfigurationConst.Configs["SBURL"].Value;
                    var pilotID = ConfigurationConst.Configs["pilotID"].Value;
                    var xmlFile = ConfigurationConst.Configs["xmlFile"].Value;

                    Utils.simbriefURL = simbriefURL;
                    Utils.pilotID = pilotID;
                    Utils.xmlFile = xmlFile;

                    msfsCommander.SBFetch();

                        
                    if (MonitoringMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "SimBrief data fetched", LOG_NORMAL);


                    msfsCommander.WASMDisconnect();


                    break;



            }

        }

        /// <summary>
        /// Looks for a VA variable to determine if additional logging should be done to the main VA window
        /// </summary>
        public static class ConfigurationConst
        {
            public static KeyValueConfigurationCollection Configs;
        }

        internal class ApplicationConfigurationReader
        {
            public void Read()
            {
                // read assembly
                var ExecAppPath = this.GetType().Assembly.Location;

                // Get all app settings  in config file
                ConfigurationConst.Configs = ConfigurationManager.OpenExeConfiguration(ExecAppPath).AppSettings.Settings;

            }
        }

        private static bool DebugMode(dynamic vaProxy)
        {
            // enables more detailed logging
            bool? result = vaProxy.GetBoolean(VARIABLE_NAMESPACE + ".DebugMode");
            if (result.HasValue)
                return result.Value;
            else
                return false;
        }

        private static bool MonitoringMode(dynamic vaProxy)
        {
            // enables more detailed logging
            bool? result = vaProxy.GetBoolean(VARIABLE_NAMESPACE + ".MonitoringMode");
            if (result.HasValue)
                return result.Value;
            else
                return false;
        }

        /// <summary>
        /// Creates a connection to the sim
        /// </summary>
        private static Commander ConnectToSim(dynamic vaProxy)
        {

            Commander msfsCommander = new Commander();
            msfsCommander.Connect();

            if (msfsCommander.Connected)
            {

                if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Successfully connected to SimConnect.", LOG_INFO);                              
                return msfsCommander;
            }
            else
            {
                if (msfsCommander != null) msfsCommander.Disconnect();
                vaProxy.WriteToLog(LOG_PREFIX + "Couldn't make a connection to SimConnect.", LOG_ERROR);
                return null;
            }



        }


        private static Commander ConnectToWASM(dynamic vaProxy)
        {

            Commander msfsCommander = new Commander();
            msfsCommander.WASMConnect1();

            if (msfsCommander.WAServerConnected)
            {

                if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Successfully connected to WASM.", LOG_INFO);

                vaProxy.SetText(VARIABLE_NAMESPACE + ".sconnect", "1");

                MSFS.Utils.errcon = false;


                return msfsCommander;
            }
            else
            {
                if (msfsCommander != null) msfsCommander.WASMDisconnect();
                vaProxy.WriteToLog(LOG_PREFIX + "Couldn't make a connection to WASM, context: " + MSFS.Utils.errvar, LOG_ERROR);

                vaProxy.SetText(VARIABLE_NAMESPACE + ".sconnect", "0");

                MSFS.Utils.errcon = true;

                return null;
            }

        }

        public static void LogOutput(string message, string color = "blank")
        {
            if (DebugMode(VA))
            {
                VA.WriteToLog("MSFS: " + message, color);
            }

        }



    }
}


