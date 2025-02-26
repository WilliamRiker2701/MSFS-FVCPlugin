//=================================================================================================================
// PROJECT: MSFS VA Plugin
// PURPOSE: This class adds Microsoft Flight Simulator 2020 full support for Voice Attack 
// AUTHOR: William Riker
//================================================================================================================= 

using System;
using System.Threading;
using System.IO;
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
using System.Collections.Generic;
using FSUIPC;
using System.Configuration;
using System.Globalization;
using Fleck;
using System.Text;
using System.Net;
using System.IO;
using System.Net.Http;
using System.Net.Sockets;
using System.Collections.Generic;
using System.Reflection;
using System.Linq;
using System.Threading.Tasks;
using System.Security.Policy;
using System.Text.RegularExpressions;
using System.Security.Cryptography;
using System.ComponentModel;
using System.Diagnostics.Metrics;
using System.Xml;

namespace MSFS
{
    public class VoiceAttackPlugin
    {
        const string VARIABLE_NAMESPACE = "MSFS";
        const string LOG_PREFIX = "MSFS: ";
        const string LOG_NORMAL = "black";
        const string LOG_ERROR = "red";
        const string LOG_INFO = "grey";
        private static dynamic VA;
        private static WebSocketServer webSocketServer;
        
        public string reqType = "NULL";
        private static IWebSocketConnection webSocketClient;
        




        public static string directoryPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

        public const string logfolderName = "Log";

        public static string logfolderPath = Path.Combine(directoryPath, logfolderName);

        public static string logfilePath = Path.Combine(logfolderPath, "log.txt");


        /// <summary>
        /// Name of the plug-in as it should be shown in the UX
        /// </summary>
        public static string VA_DisplayName()
        {
            return "MSFS-FVCplugin - v1.4.5.1-b";
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


            if (File.Exists(logfilePath))
            {
                File.Delete(logfilePath);
            }

            Directory.CreateDirectory(logfolderPath);


            VA = vaProxy;

            VA.LogEntryAdded += new Action<DateTime, String, String>(SaveLogEntry);

            webSocketServer = new WebSocketServer("ws://127.0.0.1:54325");
            

            // Start the server

            webSocketServer.Start(socket =>
            {
                // Handle the WebSocket connection
                socket.OnOpen = () =>
                {
                    VA.WriteToLog(LOG_PREFIX + "Connection with FVC Panel opened", "orange");
                    webSocketClient = socket;
                    
                };

                socket.OnClose = () =>
                {
                    VA.WriteToLog(LOG_PREFIX + "Connection with FVC Panel closed", "orange");
                    if (webSocketClient == socket)
                    {
                        webSocketClient = null;
                    }

                };

            });





            if (DebugMode(VA)) VA.WriteToLog(LOG_PREFIX + "WebSocket message: " + Utils.webSocketColor + " " + Utils.webSocketMessage, LOG_NORMAL);
        }

        /// <summary>
        /// Handles clean up before Voice Attack closes
        /// </summary>
        public static void VA_Exit1(dynamic vaProxy)
        {
            StopWebSocketServer();
        }

        /// <summary>
        /// Main function used to process commands from Voice Attack
        /// </summary>
        public static async void VA_Invoke1(dynamic vaProxy)
        {

            CultureInfo invariantCulture = CultureInfo.InvariantCulture;
            Thread.CurrentThread.CurrentCulture = invariantCulture;
            Thread.CurrentThread.CurrentUICulture = invariantCulture;



            string context;
            string targetVar;
            string eventData;
            string eventData2;
            Commander msfsCommander;
            WeatherReport msfsWeatherReport;
            string varKind = "K";
            string varAction = "Set";
            int outputVar = 0;
            int aDigit = 0;
            
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

                if (context.Substring(context.Length - 3) == "-GA")
                {
                    outputVar = 1;
                    varAction = "Get";

                }

                if (context.Substring(context.Length - 4) == "-GA2")
                {
                    outputVar = 2;
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
                if (context.Substring(context.Length - 3) == "-GA")
                {
                    outputVar = 1;
                    varAction = "Get";

                }
                if (context.Substring(context.Length - 4) == "-GA2")
                {
                    outputVar = 2;
                    varAction = "Get";

                }

            }
            
            else if (context.Substring(0, 2) == "A:")
            {

                varKind = "A";

                if (context.Substring(context.Length - 3) == "SUB")
                {
                    outputVar = 3;
                    varAction = "Norm";

                }
                if (context.Substring(context.Length - 2) == "-A")
                {
                    outputVar = 1;
                    varAction = "Norm";

                }
                if (System.Text.RegularExpressions.Regex.IsMatch(context, @"-A\d$"))
                {
                    outputVar = 2;
                    varAction = "Norm";
                    aDigit = context[context.Length - 1] - '0';

                }

                else
                {

                    varAction = "Norm";

                }

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

            else if (context.Substring(0, 3) == "FR:")
            {

                varKind = "FAC REQ";

            }

            else if (context == "SimBriefFetch")
            {

                varKind = "SBF";

            }

            else if (context == "METAR")
            {

                varKind = "METARPULL";

            }

            else if (context == "METARPARSE")
            {

                varKind = "METARPARSE";

            }

            else if (context == "CCTARGET")
            {

                varKind = "CCTARGET";

            }

            else if (context == "A310TRANS")
            {

                varKind = "A310TRANS";

            }

            else if (context == "Compact")
            {

                varKind = "Compact";

            }

            else if (context == "ServerStart")
            {

                varKind = "serverstart";

            }

            else if (context == "ServerStop")
            {

                varKind = "serverstop";

            }

            else if (context == "HEXTODEC")
            {

                varKind = "ConvertHexToDecimal";

            }

            else if (context == "Set.appSetting")
            {

                varKind = "setappSetting";

            }

            else if (context == "Read.appSetting")
            {

                varKind = "readappSetting";

            }

            else if (context == "PMDG.R")
            {

                varKind = "PMDG read serial";

            }

            else if (context == "PMDG.W")
            {

                varKind = "PMDG write serial";

            }

            else if (context == "OPACITY")
            {

                varKind = "OPACITY";

            }

            else if (context == "CLEARLOG")
            {

                varKind = "CLEARLOG";

            }

            else if (context.StartsWith("CALCULATOR"))
            {

                varKind = "Calculator";
                

            }

            switch (varKind)
            {
                case "L":

                    switch (varAction)
                    {
                        case "Set":

                            double varResult;

                            double dData;

                            int count = 1;

                            MSFS.Utils.errvar = context;

                            msfsCommander = ConnectToWASM2(vaProxy);

                            while (MSFS.Utils.errcon == true)
                            {

                                if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Retrying connection", LOG_INFO);

                                msfsCommander = ConnectToWASM2(vaProxy);

                            }
                                                                                                           
                            context = context.Remove(0, 2);

                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Processing Local Variable new value set: " + context, LOG_INFO);

                            eventData = vaProxy.GetText(VARIABLE_NAMESPACE + ".ValueSet");

                            dData = Convert.ToDouble(eventData);

                            //msfsCommander.SetLVarFSUIPC(context, eventData);
                            msfsCommander.TriggerWASM(context, eventData);

                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "SetLVarFSUIPC Executed", LOG_INFO);

                            Thread.Sleep(100);

                            //varResult = msfsCommander.GetLVarFSUIPC(context);
                            varResult = msfsCommander.TriggerWASM(context);


                            while (dData != varResult && count <= 10)
                            {
                                
                                count ++;

                                //msfsCommander.SetLVarFSUIPC(context, eventData);
                                msfsCommander.TriggerWASM(context, eventData);

                                Thread.Sleep(100);

                                //varResult = msfsCommander.GetLVarFSUIPC(context);
                                varResult = msfsCommander.TriggerWASM(context);

                                if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Set retry, New value: " + varResult, LOG_NORMAL);                                                       

                            }

                            if (MonitoringMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "New value: " + varResult + "| Variable: " + context, LOG_NORMAL);


                            msfsCommander.WASMDisconnect();                           
                           

                            break;

                        case "Target":

                            int count2 = 0;

                            MSFS.Utils.errvar = context;

                            msfsCommander = ConnectToWASM2(vaProxy);

                            while (MSFS.Utils.errcon == true)
                            {

                                if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Retrying connection", LOG_INFO);

                                msfsCommander = ConnectToWASM2(vaProxy);

                            }

                            context = context.Remove(0, 2);

                            context = context.Remove(context.Length - 2);

                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Processing Local Variable new value set: " + context, LOG_INFO);

                            targetVar = vaProxy.GetText(VARIABLE_NAMESPACE + ".TargetVar");

                            eventData2 = vaProxy.GetText(VARIABLE_NAMESPACE + ".ValueTarget");

                            eventData = vaProxy.GetText(VARIABLE_NAMESPACE + ".ValueSet");

                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Target variable: " + targetVar, LOG_INFO);

                            dData = Convert.ToDouble(eventData2);

                            //msfsCommander.SetLVarFSUIPC(context, eventData);
                            msfsCommander.TriggerWASM(context, eventData);

                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "SetLVarFSUIPC Executed", LOG_INFO);

                            Thread.Sleep(100);

                            //varResult = msfsCommander.GetLVarFSUIPC(context);
                            varResult = msfsCommander.TriggerWASM(context);
                            //varResult2 = msfsCommander.GetLVarFSUIPC("L:"+targetVar);
                            varResult2 = msfsCommander.TriggerWASM("L:" + targetVar);

                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "New Value: " + varResult + "| Variable: " + context, LOG_NORMAL);
                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Target Value: " + varResult2 + "| Target: " + targetVar, LOG_NORMAL);

                            while (dData != varResult2  && count2 < 10)
                            {
                                //msfsCommander.SetLVarFSUIPC(context, eventData);
                                msfsCommander.TriggerWASM(context, eventData);

                                Thread.Sleep(50);

                                //varResult2 = msfsCommander.GetLVarFSUIPC("L:" + targetVar);
                                varResult2 = msfsCommander.TriggerWASM("L:" + targetVar);

                                if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Set retry, New target value: " + varResult2, LOG_NORMAL);

                                count2++;

                            }





                            if (MonitoringMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "New value: " + varResult + "| Variable: " + context, LOG_NORMAL);


                            msfsCommander.WASMDisconnect();


                            break;

                        case "Get":

                            

                            string lVarSt = "NULL";

                            MSFS.Utils.errvar = context;

                            msfsCommander = ConnectToWASM2(vaProxy);

                            while (MSFS.Utils.errcon == true)
                            {

                                if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Retrying connection", LOG_INFO);

                                msfsCommander = ConnectToWASM2(vaProxy);

                            }

                            context = context.Remove(0, 2);

                            switch (outputVar)
                            {

                                case 0:
                                    context = context.Remove(context.Length - 2);
                                    break;

                                case 1:
                                    context = context.Remove(context.Length - 3);
                                    break;

                                case 2:
                                    context = context.Remove(context.Length - 4);
                                    break;
                            }


                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Processing Local Variable data request: " + context, LOG_INFO);
                                                     
                            //varResult = msfsCommander.GetLVarFSUIPC(context);
                            varResult = msfsCommander.TriggerWASM(context);

                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "varResult: " + varResult, LOG_INFO);

                            //if (msfsCommander.readsSep == true)
                            //{

                            lVarSt = varResult.ToString();

                            lVarSt = string.Join<char>(" ", lVarSt).Replace("- ", "-");

                            lVarSt = lVarSt.Replace(".", "point");

                            

                            switch (outputVar)
                            {

                                case 0:
                                    vaProxy.SetText(VARIABLE_NAMESPACE + ".ValueGetSt", lVarSt);
                                    break;

                                case 1:
                                    vaProxy.SetText(VARIABLE_NAMESPACE + ".AuxValueGetSt", lVarSt);
                                    break;

                                case 2:
                                    vaProxy.SetText(VARIABLE_NAMESPACE + ".Aux2ValueGetSt", lVarSt);
                                    break;
                            }

                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Local Variable request returns string: " + lVarSt, LOG_INFO);

                            //}

                            switch (outputVar)
                            {

                                case 0:
                                    vaProxy.SetDecimal(VARIABLE_NAMESPACE + ".ValueGet", (decimal?)(double?)varResult);
                                    break;

                                case 1:
                                    vaProxy.SetDecimal(VARIABLE_NAMESPACE + ".AuxValueGet", (decimal?)(double?)varResult);
                                    break;

                                case 2:
                                    vaProxy.SetDecimal(VARIABLE_NAMESPACE + ".Aux2ValueGet", (decimal?)(double?)varResult);
                                    break;
                            }
                                                                                                                   

                            string calcCode = "(L:" + context + ")";

                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Local Variable request returns value: " + varResult, LOG_INFO);

                            if (MonitoringMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Present value: " + varResult + "| Variable: " + context, LOG_NORMAL);
                            

                            msfsCommander.WASMDisconnect();

                            Thread.Sleep(100);

                            break;

                    }

                    break;

                case "L volume":


                    switch (varAction)
                    {
                        case "Add":

                            MSFS.Utils.errvar = context;

                            msfsCommander = ConnectToWASM2(vaProxy);

                            while (MSFS.Utils.errcon == true)
                            {

                                if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Retrying connection", LOG_INFO);

                                msfsCommander = ConnectToWASM2(vaProxy);

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

                            msfsCommander = ConnectToWASM2(vaProxy);

                            while (MSFS.Utils.errcon == true)
                            {

                                if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Retrying connection", LOG_INFO);

                                msfsCommander = ConnectToWASM2(vaProxy);

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

                            msfsCommander = ConnectToWASM2(vaProxy);

                            while (MSFS.Utils.errcon == true)
                            {

                                vaProxy.WriteToLog(LOG_PREFIX + "Retrying connection", LOG_INFO);

                                msfsCommander = ConnectToWASM2(vaProxy);

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

                            slpValue = int.Parse(thrSleep);

                            while (i < multInt)
                            {

                                msfsCommander.TriggerWASM(context, eventData);

                                i = i + 1;

                                Thread.Sleep(slpValue);

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

                            msfsCommander = ConnectToWASM2(vaProxy);

                            while (MSFS.Utils.errcon == true)
                            {

                                if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Retrying connection", LOG_INFO);

                                msfsCommander = ConnectToWASM2(vaProxy);

                            }

                            if (msfsCommander == null) return;

                            context = context.Remove(0, 2);

                            switch (outputVar)
                            {

                                case 0:
                                    context = context.Remove(context.Length - 2);
                                    break;

                                case 1:
                                    context = context.Remove(context.Length - 3);
                                    break;

                                case 2:
                                    context = context.Remove(context.Length - 4);
                                    break;
                            }


                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Processing PMDG Variable data request: " + context, LOG_INFO);

                            Thread.Sleep(50);

                            varResult = msfsCommander.TriggerReqPMDG(context);


                            switch (outputVar)
                            {

                                case 0:
                                    vaProxy.SetText(VARIABLE_NAMESPACE + ".ValueGet", varResult);
                                    break;

                                case 1:
                                    vaProxy.SetText(VARIABLE_NAMESPACE + ".AuxValueGet", varResult);
                                    break;

                                case 2:
                                    vaProxy.SetText(VARIABLE_NAMESPACE + ".Aux2ValueGet", varResult);
                                    break;
                            }

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

                    msfsCommander = ConnectToWASM2(vaProxy);

                    while (MSFS.Utils.errcon == true)
                    {

                        if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Retrying connection", LOG_INFO);

                        msfsCommander = ConnectToWASM2(vaProxy);

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

                    switch (varAction)
                    {
                        case "Norm":

                            double simvarVal;

                            string simvarValinSt = "NULL";

                            string simvarString = "NULL";

                            context = context.Remove(0, 2);

                            switch (outputVar)
                            {

                                case 0:

                                    break;

                                case 1:
                                    context = context.Remove(context.Length - 2);
                                    Utils.simVarSubscription = false;
                                    break;

                                case 2:
                                    context = context.Remove(context.Length - 3);
                                    Utils.simVarSubscription = false;
                                    break;

                                case 3:
                                    context = context.Remove(context.Length - 4);
                                    Utils.simVarSubscription = true;
                                    break;

                            }

                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Processing SimVar: " + context, LOG_INFO);

                            msfsCommander = SimvarReq(context);

                            simvarVal = Utils.resultDataValue;

                            simvarString = Utils.resultDataString;

                            if (msfsCommander.readsSep == true)
                            {

                                simvarValinSt = simvarVal.ToString();

                                simvarValinSt = string.Join<char>(" ", simvarValinSt).Replace("- ", "-");

                                simvarValinSt = simvarValinSt.Replace(".", "point");

                                switch (outputVar)
                                {

                                    case 0:
                                        vaProxy.SetText(VARIABLE_NAMESPACE + ".ValueGetSt", simvarValinSt);
                                        break;

                                    case 1:
                                        vaProxy.SetText(VARIABLE_NAMESPACE + ".AuxValueGetSt", simvarValinSt);
                                        break;

                                    case 2:
                                        vaProxy.SetText(VARIABLE_NAMESPACE + ".Aux" + aDigit + "ValueGetSt", simvarValinSt);
                                        break;

                                }

                                if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "SimVar with separations: " + simvarValinSt, LOG_NORMAL);

                            }

                            switch (outputVar)
                            {

                                case 0:
                                    vaProxy.SetDecimal(VARIABLE_NAMESPACE + ".ValueGet", (decimal?)(double?)simvarVal);
                                    vaProxy.SetText(VARIABLE_NAMESPACE + ".StringGet", simvarString);
                                    break;

                                case 1:
                                    vaProxy.SetDecimal(VARIABLE_NAMESPACE + ".AuxValueGet", (decimal?)(double?)simvarVal);
                                    vaProxy.SetText(VARIABLE_NAMESPACE + ".AuxStringGet", simvarString);
                                    break;

                                case 2:
                                    vaProxy.SetDecimal(VARIABLE_NAMESPACE + ".Aux" + aDigit + "ValueGet", (decimal?)(double?)simvarVal);
                                    vaProxy.SetText(VARIABLE_NAMESPACE + ".Aux" + aDigit + "StringGet", simvarString);
                                    break;

                            }


                            

                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "SimVar request returns value: " + simvarVal, LOG_NORMAL);
                            if (DebugMode(vaProxy))
                            {
                                if (Utils.resultDataString != "NULL") vaProxy.WriteToLog(LOG_PREFIX + "SimVar request returns string: " + simvarString, LOG_NORMAL);
                            }
                            if (MonitoringMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Present value: " + simvarVal + "| Variable: " + context, LOG_NORMAL);
                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Present value: " + simvarVal + "| Variable: " + context, LOG_NORMAL);

                            break;


                    }

                    break;

                case "CalcCode":

                    double varRes;

                    MSFS.Utils.errvar = context;


                    msfsCommander = ConnectToWASM2(vaProxy);

                    while (MSFS.Utils.errcon == true)
                    {

                        if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Retrying connection", LOG_INFO);

                        msfsCommander = ConnectToWASM2(vaProxy);

                    }


                    context = context.Remove(0, 3);


                    if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Sending calculator code: " + context, LOG_INFO);

                    varRes = msfsCommander.TriggerCalcCode(context);

                    vaProxy.SetDecimal(VARIABLE_NAMESPACE + ".ValueGet", (decimal?)(double?)varRes);

                    vaProxy.SetText(VARIABLE_NAMESPACE + ".StringGet", Utils.calcString);

                    if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Calculator code response: " + varRes, LOG_NORMAL);

                    msfsCommander.WASMDisconnect();

                    Thread.Sleep(100);

                    break;

                case "FAC REQ":


                    msfsCommander = ConnectToWASM2(vaProxy);

                    while (MSFS.Utils.errcon == true)
                    {

                        vaProxy.WriteToLog(LOG_PREFIX + "Retrying connection", LOG_INFO);

                        msfsCommander = ConnectToWASM2(vaProxy);

                    }

                    if (msfsCommander == null) return;

                    if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Processing SimConnect Facility Data request...", LOG_INFO);

                    context = context.Remove(0, 3);

                    VoiceAttackPlugin.FacilityReq(context);

                    msfsCommander.FacilityWrappingUp();

                    msfsCommander.Disconnect();

                    msfsCommander.WASMDisconnect();

                    Thread.Sleep(100);

                    break;

                case "SBF":


                    msfsCommander = ConnectToWASM2(vaProxy);

                    while (MSFS.Utils.errcon == true)
                    {

                        vaProxy.WriteToLog(LOG_PREFIX + "Retrying connection", LOG_INFO);

                        msfsCommander = ConnectToWASM2(vaProxy);

                    }

                    if (msfsCommander == null) return;

                    if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Processing SimBrief request...", LOG_INFO);

                    new ApplicationConfigurationReader().Read();
                    string simbriefURL = "https://www.simbrief.com/api/xml.fetcher.php?userid={0}";
                    var pilotID = ConfigurationConst.Configs["pilotID"].Value;

                    Utils.simbriefURL = simbriefURL;
                    Utils.pilotID = pilotID;

                    msfsCommander.SBFetch();

                        
                    if (MonitoringMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "SimBrief data fetched", LOG_NORMAL);


                    msfsCommander.WASMDisconnect();

                    Thread.Sleep(100);


                    break;

                case "METARPULL":

                    string ICAO = vaProxy.GetText(VARIABLE_NAMESPACE + ".ValueSet");

                    ICAO = ICAO.ToUpper();

                    string urlMetar = "https://aviationweather.gov/cgi-bin/data/metar.php?ids=" + ICAO + "&hours=0&format=raw";
                    string urlTaf = "https://aviationweather.gov/cgi-bin/data/taf.php?ids=" + ICAO + "&hours=0&format=raw";

                    string metar = GetPlainTextFromUrl(urlMetar).GetAwaiter().GetResult();
                    string taf = GetPlainTextFromUrl(urlTaf).GetAwaiter().GetResult();

                    if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "METAR report pulled: " + metar, LOG_INFO);
                    if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "TAF report pulled: " + taf, LOG_INFO);

                    if (metar.Length>1)
                        metar = metar.Substring(0, metar.Length - 1);
                    if (taf.Length > 1)
                        taf = taf.Substring(0, taf.Length - 1);

                    vaProxy.SetText(VARIABLE_NAMESPACE + ".Metar", metar);
                    vaProxy.SetText(VARIABLE_NAMESPACE + ".Taf", taf);


                    break;

                case "METARPARSE":
                    
                    WeatherReport weatherReport = new WeatherReport();
                    string reportType;

                    if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "METAR parse called.", LOG_INFO);
                                        
                    try
                    {                        
                        string metarRep = vaProxy.GetText(VARIABLE_NAMESPACE + ".ValueSet");

                        if (Regex.IsMatch(metarRep, @"\b\d{4}/\d{4}\b"))
                        {
                            reportType = "TAF";
                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "TAF report: " + metarRep, LOG_INFO);
                        }
                        else
                        {
                            reportType = "METAR";
                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "METAR report: " + metarRep, LOG_INFO);
                        }                                                                      

                       if (reportType == "METAR")
                        {
                            weatherReport.MetarReport(metarRep);

                            // if RMK another function
                            if (Utils.tempoRepExists) 
                            {
                                weatherReport.AltSection(Utils.tempoRep);
                                if (Utils.becmgRepExists)
                                {
                                    weatherReport.AltSection(Utils.becmgRep);
                                }
                            }
                            if (Utils.becmgRepExists)
                            {
                                weatherReport.AltSection(Utils.becmgRep);
                                if (Utils.tempoRepExists)
                                {
                                    weatherReport.AltSection(Utils.tempoRep);
                                }
                            }

                        }
                        else
                        {
                            weatherReport.TafReport(metarRep);
                        }

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error parsing METAR report: {ex.Message}");
                    }

                    break;

                case "CCTARGET":


                    MSFS.Utils.errvar = context;


                    msfsCommander = ConnectToWASM2(vaProxy);

                    while (MSFS.Utils.errcon == true)
                    {

                        if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Retrying connection", LOG_INFO);

                        msfsCommander = ConnectToWASM2(vaProxy);

                    }

                    if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Method invoked CCTARGET", LOG_INFO);

                    string upCode;
                    string downCode;
                    double modelModValue;

                    string modelVariable = vaProxy.GetText(VARIABLE_NAMESPACE + ".CCMODEL");
                    string targetVariable = vaProxy.GetText(VARIABLE_NAMESPACE + ".CCTARGET");
                    string methodType = vaProxy.GetText(VARIABLE_NAMESPACE + ".CCTYPE");
                    string modelModification = vaProxy.GetText(VARIABLE_NAMESPACE + ".CCMODELMOD");
                    
                    int threadSleep = vaProxy.GetInt(VARIABLE_NAMESPACE + ".CCTHREADSLEEP");

                    double modelValue = msfsCommander.TriggerCalcCode(modelVariable);
                    double targetValue = msfsCommander.TriggerCalcCode(targetVariable);

                    

                    if (modelModification != "NO")
                    {                        
                        switch (modelModification)
                        {
                            case "DIV":

                                modelModValue = Convert.ToDouble(vaProxy.GetText(VARIABLE_NAMESPACE + ".CCMODELMODVALUE"));

                                modelValue = modelValue / modelModValue;

                            break;

                            case "MULT":

                                modelModValue = Convert.ToDouble(vaProxy.GetText(VARIABLE_NAMESPACE + ".CCMODELMODVALUE"));

                                modelValue = modelValue * modelModValue;

                            break;
                        }
                    }

                    if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Model: " + modelValue + " and target value: " + targetValue, LOG_INFO);
                    if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Method type: " + methodType + " and model modification type: " + modelModification, LOG_INFO);


                    switch (methodType)
                    {
                        case "UPDOWN":

                            upCode = vaProxy.GetText(VARIABLE_NAMESPACE + ".CCUPCODE");
                            downCode = vaProxy.GetText(VARIABLE_NAMESPACE + ".CCDOWNCODE");

                            while (targetValue < modelValue)
                            {
                                msfsCommander.TriggerCalcCode(upCode);
                                Thread.Sleep(threadSleep);
                                targetValue = msfsCommander.TriggerCalcCode(targetVariable);

                            }

                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Target value after upCode: " + targetValue, LOG_INFO);

                            while (targetValue > modelValue)
                            {
                                msfsCommander.TriggerCalcCode(downCode);
                                Thread.Sleep(threadSleep);
                                targetValue = msfsCommander.TriggerCalcCode(targetVariable);

                            }

                            if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Target value after upCode and downCode: " + targetValue, LOG_INFO);

                            break;

                    }
                    
                    msfsCommander.WASMDisconnect();

                    Thread.Sleep(100);

                    break;

                case "A310TRANS":


                    MSFS.Utils.errvar = context;


                    msfsCommander = ConnectToWASM2(vaProxy);

                    while (MSFS.Utils.errcon == true)
                    {

                        if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Retrying connection", LOG_INFO);

                        msfsCommander = ConnectToWASM2(vaProxy);

                    }

                    int xpndr1 = int.Parse(vaProxy.GetText(VARIABLE_NAMESPACE + ".XPNDR1"));
                    int xpndr2 = int.Parse(vaProxy.GetText(VARIABLE_NAMESPACE + ".XPNDR2"));

                    if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Transponder target value: " + xpndr1 + xpndr2, LOG_INFO);

                    msfsCommander.TriggerCalcCode("(A:TRANSPONDER CODE:1, Number)");

                    int xpndr3 = int.Parse(Utils.calcString.Substring(0, 2));
                    int xpndr4 = int.Parse(Utils.calcString.Substring(2, 2));

                    if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Transponder current value: " + xpndr3 + xpndr4, LOG_INFO);

                    while (xpndr3 < xpndr1)
                    {
                        msfsCommander.TriggerCalcCode("1 (>L:A310_TRANSPONDER_OUTER_KNOB_TURNED_CLOCKWISE)");
                        Thread.Sleep(50);
                        msfsCommander.TriggerCalcCode("(A:TRANSPONDER CODE:1, Number)");
                        xpndr3 = int.Parse(Utils.calcString.Substring(0, 2));
                    }
                    while (xpndr3 > xpndr1)
                    {
                        msfsCommander.TriggerCalcCode("1 (>L:A310_TRANSPONDER_OUTER_KNOB_TURNED_ANTICLOCKWISE)");
                        Thread.Sleep(50);
                        msfsCommander.TriggerCalcCode("(A:TRANSPONDER CODE:1, Number)");
                        xpndr3 = int.Parse(Utils.calcString.Substring(0, 2));
                    }

                    while (xpndr4 < xpndr2)
                    {
                        msfsCommander.TriggerCalcCode("1 (>L:A310_TRANSPONDER_INNER_KNOB_TURNED_CLOCKWISE)");
                        Thread.Sleep(50);
                        msfsCommander.TriggerCalcCode("(A:TRANSPONDER CODE:1, Number)");
                        xpndr4 = int.Parse(Utils.calcString.Substring(2, 2));
                    }
                    while (xpndr4 > xpndr2)
                    {
                        msfsCommander.TriggerCalcCode("1 (>L:A310_TRANSPONDER_INNER_KNOB_TURNED_ANTICLOCKWISE)");
                        Thread.Sleep(50);
                        msfsCommander.TriggerCalcCode("(A:TRANSPONDER CODE:1, Number)");
                        xpndr4 = int.Parse(Utils.calcString.Substring(2, 2));
                    }

                    msfsCommander.WASMDisconnect();

                    Thread.Sleep(100);

                    break;

                case "Compact":

                    if (vaProxy.GetBoolean("SetCompact") == true)
                    {
                        if (vaProxy.CompactModeEnabled == false)
                        {
                            vaProxy.ToggleCompactMode();
                            vaProxy.WriteToLog("Compact mode on", "black");
                        }
                        else
                        {
                            vaProxy.WriteToLog("Compact mode already on", "black");
                        }
                    }
                    else
                    {
                        if (vaProxy.CompactModeEnabled == true)
                        {
                            vaProxy.ToggleCompactMode();
                            vaProxy.WriteToLog("Compact mode off", "black");
                        }
                        else
                        {
                            vaProxy.WriteToLog("Compact mode already off", "black");
                        }
                    }
                    

                    break;

                case "serverstart":

                    StartWebSocketServer();

                    break;

                case "serverstop":

                    StopWebSocketServer();

                    break;

                case "ConvertHexToDecimal":

                    string hexNumber = vaProxy.GetText("HexNumber");

                    int decimalValue = 0;

                    for (int j = hexNumber.Length - 1, power = 0; j >= 0; j--, power++)
                    {
                        int digitValue;

                        if (char.IsDigit(hexNumber[j]))
                        {
                            digitValue = hexNumber[j] - '0';
                        }
                        else
                        {
                            digitValue = char.ToUpper(hexNumber[j]) - 'A' + 10;
                        }

                        decimalValue += digitValue * (int)Math.Pow(16, power);
                    }

                    vaProxy.SetDecimal(VARIABLE_NAMESPACE + ".ValueGet", (decimal?)(double?)decimalValue);

                    break;


                case "setappSetting":

                    if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Processing app.Setting write request...", LOG_INFO);

                    eventData = vaProxy.GetText("appSetting.ValueSet");
                    eventData2 = vaProxy.GetText("appSetting");

                    new ApplicationConfigurationReader().Overwrite(eventData2,eventData);

                    break;

                case "readappSetting":

                    if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Processing app.Setting read request...", LOG_INFO);

                    eventData = vaProxy.GetText("appSetting");

                    try
                    {
                        new ApplicationConfigurationReader().Read();
                        var result = ConfigurationConst.Configs[eventData].Value;
                        vaProxy.SetText("appSetting.ValueGet", result);
                    }
                    catch
                    {
                        if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "appSetting NOT SET", LOG_INFO);
                        vaProxy.SetText("appSetting.ValueGet", "Not Set");
                    }



                    break;

                case "PMDG read serial":

                    string serial = "";
                    string value;
                    string zeroLetters = "acegikmoqsuwy";
                    string oneLetters = "bdfhjlnprtvxz";
                    string baseDirectory = vaProxy.PluginPath();
                    string filePath = Path.Combine(baseDirectory.Substring(0, baseDirectory.Length - 18), "System.Value.dat");
                    vaProxy.SetText("FirstTimePMDGERROR", "0");

                    if (File.Exists(filePath))
                    {
                        try
                        {
                            serial = File.ReadAllText(filePath);

                            int length = serial.Length;

                            if (length == 16)
                            {
                                if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Serial length correct. Processing serial.", LOG_INFO);

                                char seventhChar = serial[6];

                                string seventh = seventhChar.ToString();

                                bool isPresent = zeroLetters.Contains(seventhChar);

                                value = isPresent.ToString();

                                vaProxy.SetText("FirstTimePMDG", value);
                            }

                            else
                            {
                                if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Serial length incorrect. Assuming first run.", LOG_INFO);
                                vaProxy.SetText("FirstTimePMDG", "True");
                            }

                        }
                        catch (Exception ex)
                        {
                            vaProxy.WriteToLog("PMDG init error: " + ex, "red");
                            vaProxy.SetText("FirstTimePMDGERROR", "1");
                            serial = "";
                        }
                    }
                    else
                    {
                        if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Serial file not found. Assuming first run.", LOG_INFO);
                        vaProxy.SetText("FirstTimePMDG", "True");
                    }
                                                   

                    break;

                case "PMDG write serial":

                    oneLetters = "bdfhjlnprtvxz";

                    baseDirectory = vaProxy.PluginPath();
                    filePath = Path.Combine(baseDirectory.Substring(0, baseDirectory.Length - 18), "System.Value.dat");
                    vaProxy.SetText("FirstTimePMDGERROR", "0");
                    Random random = new Random();

                    serial = "";

                    int randomIndex = random.Next(oneLetters.Length);
                    char randomLetter = oneLetters[randomIndex];

                    // Add some dummy letters and numbers
                    serial += GetRandomString(3); // 3 random characters
                    serial += randomLetter;
                    serial += GetRandomString(12); // 12 random characters

                    try
                    {
                        File.WriteAllText(filePath, serial);
                    }
                    catch (Exception ex)
                    {
                        vaProxy.WriteToLog("PMDG init error type W: " + ex, "red");

                    }

                    break;

                case "Calculator":

                    Utils.Calculator(context.Substring(11));

                    break;

                case "OPACITY":

                    int opacity = vaProxy.GetInt("VAopacity");

                    vaProxy.SetOpacity(opacity);

                    break;

                case "CLEARLOG":

                    vaProxy.ClearLog();

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

                if (DebugMode(VA)) VA.WriteToLog(LOG_PREFIX + "Reading app.Setting.", LOG_INFO);

                // read assembly
                var ExecAppPath = this.GetType().Assembly.Location;

                // Get all app settings  in config file
                ConfigurationConst.Configs = ConfigurationManager.OpenExeConfiguration(ExecAppPath).AppSettings.Settings;

                if (DebugMode(VA)) VA.WriteToLog(LOG_PREFIX + "Reading app.Setting done.", LOG_INFO);

            }

            public void Overwrite(string key, string value)
            {

                if (DebugMode(VA)) VA.WriteToLog(LOG_PREFIX + "Overwriting app.Setting.", LOG_INFO);

                var ExecAppPath = this.GetType().Assembly.Location;

                Configuration config = ConfigurationManager.OpenExeConfiguration(ExecAppPath);
                config.AppSettings.Settings.Remove(key);
                config.AppSettings.Settings.Add(key, value);
                config.Save(ConfigurationSaveMode.Minimal);
                ConfigurationManager.RefreshSection("appSettings");

                if (DebugMode(VA)) VA.WriteToLog(LOG_PREFIX + "Overwriting app.Setting done.", LOG_INFO);

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


        private static Commander FacilityReq(string type)
        {

            Commander msfsCommander = new Commander();

            msfsCommander.DefineRequestType(type);

            Utils.facilityType = type;

            string reqRunway;

            switch (type) 
            {
                case "AIRPORT":

                    Utils.reqICAO = VA.GetText(VARIABLE_NAMESPACE + ".FacilityReqICAO");

                    if (MonitoringMode(VA)) VA.WriteToLog(LOG_PREFIX + "Facility AIRPORT data requested from " + Utils.reqICAO, LOG_INFO);

                    msfsCommander.FacilityRequest();

                    break;

                case "RUNWAY":

                    Utils.reqICAO = VA.GetText(VARIABLE_NAMESPACE + ".FacilityReqICAO");
                    reqRunway = VA.GetText("MSFS.FacilityReqRunway");

                    if (reqRunway.Substring(0, 1) == "0")
                    {
                        reqRunway = reqRunway.Remove(0, 1);
                    }
                        
                    Utils.reqRunway = reqRunway;
                    Utils.simConnectMSGLoop = 1;

                    if (MonitoringMode(VA)) VA.WriteToLog(LOG_PREFIX + "Facility RUNWAY data requested from " + Utils.reqICAO + " " + Utils.reqRunway + ".", LOG_INFO);

                    msfsCommander.FacilityRequest();

                    break;

                case "VOR":

                    Utils.reqVORICAO = VA.GetText(VARIABLE_NAMESPACE + ".FacilityReqVORICAO");
                    Utils.reqVORregion = VA.GetText("MSFS.FacilityReqVORregion");
                    Utils.scPrimVORICAO1 = "NULL";
                    Utils.scPrimVORregion1 = "NULL";
                    Utils.scPrimVORICAO2 = "NULL";
                    Utils.scPrimVORregion2 = "NULL";
                    Utils.scPrimVORICAO3 = "NULL";
                    Utils.scPrimVORregion3 = "NULL";
                    Utils.scPrimVORICAO4 = "NULL";
                    Utils.scPrimVORregion4 = "NULL";
                    Utils.scPrimVORICAO5 = "NULL";
                    Utils.scPrimVORregion5 = "NULL";
                    Utils.scPrimVORICAO6 = "NULL";
                    Utils.scPrimVORregion6 = "NULL";
                    Utils.scPrimVORICAO7 = "NULL";
                    Utils.scPrimVORregion7 = "NULL";
                    Utils.scPrimVORICAO8 = "NULL";
                    Utils.scPrimVORregion8 = "NULL";

                    Utils.scSecVORICAO1 = "NULL";
                    Utils.scSecVORregion1 = "NULL";
                    Utils.scSecVORICAO2 = "NULL";
                    Utils.scSecVORregion2 = "NULL";
                    Utils.scSecVORICAO3 = "NULL";
                    Utils.scSecVORregion3 = "NULL";
                    Utils.scSecVORICAO4 = "NULL";
                    Utils.scSecVORregion4 = "NULL";
                    Utils.scSecVORICAO5 = "NULL";
                    Utils.scSecVORregion5 = "NULL";
                    Utils.scSecVORICAO6 = "NULL";
                    Utils.scSecVORregion6 = "NULL";
                    Utils.scSecVORICAO7 = "NULL";
                    Utils.scSecVORregion7 = "NULL";
                    Utils.scSecVORICAO8 = "NULL";
                    Utils.scSecVORregion8 = "NULL";

                    if (Utils.reqVORICAO == "" || Utils.reqVORICAO == "NULL")
                    {
                        if (MonitoringMode(VA)) VA.WriteToLog(LOG_PREFIX + "No data for facility VOR request", LOG_ERROR);
                    }
                    else
                    {
                        if (MonitoringMode(VA)) VA.WriteToLog(LOG_PREFIX + "Facility VOR data requested from " + Utils.reqVORICAO + " Region " + Utils.reqVORregion, LOG_INFO);

                        msfsCommander.FacilityRequest();

                    }

                    break;
            }
            


            return msfsCommander;

        }

        private static Commander SimvarReq(string simvar)
        {

            Commander msfsCommander = new Commander();

            if (MonitoringMode(VA)) VA.WriteToLog(LOG_PREFIX + "SimVar value requested: " + simvar, LOG_INFO);
            
            msfsCommander.GetSimVarSimConnect_1(simvar);

            msfsCommander.DisconnectSimConnect();

            return msfsCommander;

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

        private static Commander ConnectToWASM2(dynamic vaProxy)
        {

            Commander msfsCommander = new Commander();
            
            
            try
            {
                msfsCommander.WASMConnect3();

                if (MSFS.Utils.errcon == false)
                {
                    if (DebugMode(vaProxy)) vaProxy.WriteToLog(LOG_PREFIX + "Successfully connected to WASM.", LOG_INFO);
                }                

                


                return msfsCommander;
            }
            catch(Exception ex) 
            {
                if (msfsCommander != null) msfsCommander.WASMDisconnect();
                vaProxy.WriteToLog(LOG_PREFIX + "Couldn't make a connection to WASM, context: " + MSFS.Utils.errvar, LOG_ERROR);
                vaProxy.WriteToLog(LOG_PREFIX + "Error: " + ex, LOG_ERROR);

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

        public static void LogMonitorOutput(string message, string color = "blank")
        {
            if (MonitoringMode(VA))
            {
                VA.WriteToLog("MSFS: " + message, color);
            }

        }

        public static void ForceLogOutput(string message, string color = "blank")
        {
            
            VA.WriteToLog("MSFS: " + message, color);

        }

        public static void LogErrorOutput(string message, string color = "blank")
        {

            VA.WriteToLog("FVC Plugin Error. " + message, color);

        }

        public static void SetText(string vaVar, string cVar)
        {

            VA.SetText(vaVar, cVar);

        }
        public static void SetInt(string vaVar, int cVar)
        {

            VA.SetInt(vaVar, cVar);

        }

        public static void SetDecimal(string vaVar, decimal cVar)
        {

            VA.SetDecimal(vaVar, cVar);

        }

        public static string GetText(string vaVar)
        {

            return VA.GetText(vaVar);

        }

        public static int GetInt(string vaVar)
        {

            return VA.GetInt(vaVar);

        }

        private static void StartWebSocketServer()
        {
            try
            {

                VA.WriteToLog(LOG_PREFIX + "Starting Web Socket", LOG_NORMAL);

                int startPort = 40000; // Start outside the dynamic range
                int endPort = 49151;   // End just before the dynamic range
                bool isBound = false;

                for (int port = startPort; port <= endPort && !isBound; port++)
                {
                    try
                    {
                        var server = new WebSocketServer($"ws://0.0.0.0:{port}");
                        server.Start(socket =>
                        {
                            socket.OnOpen = () => Console.WriteLine("WebSocket Opened");
                            socket.OnClose = () => Console.WriteLine("WebSocket Closed");
                        });
                        VA.WriteToLog(LOG_PREFIX + $"Server bound to port {port}", LOG_NORMAL);
                        isBound = true;
                    }
                    catch (Exception ex)
                    {
                        VA.WriteToLog(LOG_PREFIX + $"Failed to bind to port {port}: {ex.Message}", LOG_NORMAL);
                    }
                }

                if (!isBound)
                {
                    VA.WriteToLog(LOG_PREFIX + "Could not bind to any port in the range!", LOG_NORMAL);
                }


                if (MonitoringMode(VA)) VA.WriteToLog(LOG_PREFIX + "WebSocket server started on ws://127.0.0.1:54325", LOG_NORMAL);

                if (DebugMode(VA)) VA.WriteToLog(LOG_PREFIX + "WebSocket message: " + Utils.webSocketColor + " " + Utils.webSocketMessage, LOG_NORMAL);


                StopWebSocketServer();

            }
            catch (Exception ex)
            {
                if (MonitoringMode(VA)) VA.WriteToLog(LOG_PREFIX + "Error starting WebSocket server: " + ex, LOG_ERROR);
            }

        }

        private static void StopWebSocketServer()
        {
            try
            {
                // Stop the WebSocket server
                webSocketServer?.Dispose();
                
                if (MonitoringMode(VA)) VA.WriteToLog(LOG_PREFIX + "WebSocket panel server stopped", LOG_NORMAL);
            }
            catch (Exception ex)
            {
                if (MonitoringMode(VA)) VA.WriteToLog(LOG_PREFIX + "Error stopping WebSocket server: " + ex, LOG_ERROR);
            }

        }


        public static void SaveLogEntry(DateTime theDate, String theMessage, String theIconColor)
        {


            using (StreamWriter writer = File.AppendText(logfilePath))
            {
                // Add the content as a new paragraph
                writer.WriteLine(theMessage);
            }
            

            if (webSocketClient != null)
            {
                webSocketClient.Send(theIconColor + " " + theMessage);
            }
        }

        public static string GetRandomString(int length)
        {
            Random random = new Random();
            const string chars = "abcdefghijklmnopqrstuvwxyz0123456789";
            return new string(Enumerable.Repeat(chars, length)
            .Select(s => s[random.Next(s.Length)]).ToArray());
        }


        public static async Task<string> GetPlainTextFromUrl(string url)
        {
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    string result = await client.GetStringAsync(url);
                    return result;
                }
            }
            catch (Exception ex)
            {
                VA.WriteToLog(LOG_PREFIX + "URL text error: " + ex, "red");
                return null;
            }
        }



    }
}


