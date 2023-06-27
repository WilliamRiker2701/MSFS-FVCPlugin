//=================================================================================================================
// PROJECT: MSFS VAPlugin
// PURPOSE: This class does the interfacing to SimConnect and WASM for Microsoft Flight Simulator 2020
// AUTHOR: William Riker
//================================================================================================================= 
using System;
using System.Collections.Generic;
using System.Threading;
using System.Diagnostics;
using Microsoft.FlightSimulator.SimConnect;
using FSUIPC;
using WASimCommander.CLI.Enums;
using WASimCommander.CLI.Structs;
using WASimCommander.Client;
using WASimCommander.Enums;
using WASimCommander.CLI.Client;
using System.Runtime.Remoting.Contexts;
using System.Security.Cryptography.X509Certificates;
using System.CodeDom;
using System.Security.Cryptography;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Net;




namespace MSFS
{
    /// Main class that operates against the flight simulator
    public class Commander
    {
        Microsoft.FlightSimulator.SimConnect.SimConnect _simConnection;
        WASimCommander.CLI.Client.WASimClient _waSimConnection;

        const int WM_USER_SIMCONNECT = 0x402;
        bool _connected = false;
        bool _wasmconnected = false;
        public bool readsSep = false;
        public string current = "NULL";


        public enum WasmModuleStatus
        {
            Unknown, NotFound, Found, Connected
        }

        public bool WasmAvailable => WasmStatus != WasmModuleStatus.NotFound;
        public WasmModuleStatus WasmStatus { get; private set; } = WasmModuleStatus.Unknown;


        Dictionary<RequestTypes, bool> requestPending = new Dictionary<RequestTypes, bool>();
        PlaneState _planeState = new PlaneState();

        // used for polling for messages from the sim
        EventWaitHandle _simConnectEventHandle = new EventWaitHandle(false, EventResetMode.AutoReset);
        Thread _simConnectReceiveThread = null;

        /// <summary>
        /// Returns the last known state of the plane
        /// </summary>
        public PlaneState GetPlaneState
        {
            get => _planeState;
            private set
            {
                _planeState = value;
            }
        }

        /// <summary>
        /// Checks if a Request is still waiting for a response from the sim
        /// </summary>
        public bool RequestPending(RequestTypes requestType)
        {
            return requestPending[requestType];
        }

        /// <summary>
        /// Indicates if there is a connection to the sim
        /// </summary>
        public bool Connected

        {
            get => _connected;
            private set
            {
                if (_connected != value)
                {
                    _connected = value;
                }
            }
        }

        public bool WAServerConnected

        {
            get => _wasmconnected;
            private set
            {
                if (_wasmconnected != value)
                {
                    _wasmconnected = value;
                }
            }
        }



        /// <summary>
        /// Initiates a connection the sim and initializes the agent
        /// </summary>
        public void Connect()
        {
            if (_simConnection == null)
            {
                try
                {
                    _simConnection = new Microsoft.FlightSimulator.SimConnect.SimConnect("msfs-agent", IntPtr.Zero, WM_USER_SIMCONNECT, null, 0);
                    initEventHandlers();
                    initEvents();
                    Connected = true;

      


                }
                catch (Exception ex)
                {
                    VoiceAttackPlugin.LogOutput("Unable to connect to sim.", "grey");
                }
            }

        }


        public void WASMConnect()
        {

            if (_waSimConnection == null)
            {

                try
                {
                    _waSimConnection = new WASimCommander.CLI.Client.WASimClient(0xC57E57E9);
                    initWASMHandlers();
                    _waSimConnection.connectSimulator(2000U);
                    _waSimConnection.connectServer();
                    WAServerConnected = _waSimConnection.isConnected();

                }
                catch (Exception ex)
                {
                    VoiceAttackPlugin.LogOutput("Unable to connect to WASM.", "grey");
                }
            }



        }

        public void WASMConnect1()
        {

            if (_waSimConnection == null)
            {

                try
                {
                    _waSimConnection = new WASimCommander.CLI.Client.WASimClient(0xC57E57E9);
                    _waSimConnection.connectSimulator();
                    _waSimConnection.connectServer();
                    WAServerConnected = _waSimConnection.isConnected();

                }
                catch (Exception ex)
                {
                    VoiceAttackPlugin.LogOutput("Unable to connect to WASM.", "grey");
                }
            }



        }

        public void WASMConnect2()
        {

            if (_waSimConnection == null)
            {

                VoiceAttackPlugin.LogOutput("WASM client not initialized.", "grey");

            }

            HR hr;
            int count = 0;

            do
            {
                _waSimConnection = new WASimCommander.CLI.Client.WASimClient(0xC57E57E9);
                hr = _waSimConnection.connectSimulator(2000U);
            }
            while (hr == HR.TIMEOUT && ++count < 11);

            if (hr == HR.OK) 
            {

                initWASMHandlers();
                hr = _waSimConnection.connectServer();
            }


                
            else

                VoiceAttackPlugin.LogOutput("WASM Client could not connect to SimConnect for unknown reason.", "grey");

            if (hr != HR.OK)
            {
                WasmStatus = WasmModuleStatus.NotFound;
                VoiceAttackPlugin.LogOutput("WASM Server not found or couldn't connect.", "grey");
                return;
            }

            WasmStatus = WasmModuleStatus.Connected;
            WAServerConnected = _waSimConnection.isConnected();
            VoiceAttackPlugin.LogOutput("Connected to WASimConnect Server.", "grey");


        }





        public void WASMConnectlongtimeout()
        {

            if (_waSimConnection == null)
            {
                try
                {
                    _waSimConnection = new WASimCommander.CLI.Client.WASimClient(0xC57E57E9);
                    initWASMHandlers();
                    _waSimConnection.connectServer(2000);
                    WAServerConnected = _waSimConnection.isConnected();

                }
                catch (Exception ex)
                {
                    VoiceAttackPlugin.LogOutput("Unable to connect to WASM.", "grey");
                }
            }



        }



        /// <summary>
        /// Starts a background thread to periodically check for messages back from the sim
        /// </summary>
        public void EnableMessagePolling()
        {
            _simConnectReceiveThread = new Thread(new ThreadStart(SimConnect_MessageReceiveThreadHandler));
            _simConnectReceiveThread.IsBackground = true;
            _simConnectReceiveThread.Start();
        }

        /// <summary>
        /// Turns of background thread used message polling with the sim
        /// </summary>
        public void DisableMessagePolling()
        {

            if (_simConnectReceiveThread != null)
            {
                _simConnectReceiveThread.Abort();
                _simConnectReceiveThread.Join();
                _simConnectReceiveThread = null;
            }

        }

        /// <summary>
        /// Disconnects the agent from the sim and cleans up.
        /// </summary>
        public void Disconnect()
        {

            if (!Connected) return;

            try
            {
                DisableMessagePolling();

                _simConnection.UnsubscribeFromSystemEvent(FsControlList.SIMSTART);
                _simConnection.UnsubscribeFromSystemEvent(FsControlList.SIMSTOP);
                _simConnection.UnsubscribeFromSystemEvent(FsControlList.PAUSE);
                _simConnection.Dispose();

                VoiceAttackPlugin.LogOutput("Connection to sim closed.", "grey");
            }
            catch (Exception ex)
            {
                VoiceAttackPlugin.LogOutput("Failed to Disconnect and clean up.", "grey");
            }
            finally
            {
                _simConnectReceiveThread = null;
                _simConnection = null;
                _simConnectReceiveThread = null;
                Connected = false;
            }
        }



        private void initEvents()
        {

            string keyName = Utils.current;

            Enum.TryParse(keyName, out FsControlList item);

            {
                _simConnection.MapClientEventToSimEvent(item, keyName);
                _simConnection.AddClientEventToNotificationGroup(NOTIFICATION_GROUPS.DEFAULT, item, false);
            }

        }


        public void WASMDisconnect()
        {

            if (!WAServerConnected) return;

            try
            {

                _waSimConnection.disconnectServer();
                _waSimConnection.disconnectSimulator();
                // delete the client
                _waSimConnection.Dispose();

                VoiceAttackPlugin.LogOutput("Connection to WASM closed.", "grey");
            }
            catch (Exception ex)
            {
                VoiceAttackPlugin.LogOutput("Failed to Disconnect WASM.", "grey");
            }
            finally
            {

                _waSimConnection = null;
                WAServerConnected = false;
            }

        }



        #region Operations

        /// <summary>
        /// Displays a "tip" message windows in the sim for the duration given (in seconds)
        /// </summary>
        /// <param name="text"></param>
        /// <param name="duration"></param>
        public void SetText(string text, int duration = 3)
        {

            //colors don't seem to be supported in msfs
            _simConnection.Text(SIMCONNECT_TEXT_TYPE.PRINT_WHITE, duration, null, text);
        }



        /// <summary>
        /// Set variable through WASM
        /// </summary>
        /// <param name="varName"></param>
        /// <param name="varData"></param>
        public void TriggerWASM(string varName, string varData = "0")
        {
            VoiceAttackPlugin.LogOutput("Triggering WASM write method...", "grey");

            double dData;

            if (String.IsNullOrWhiteSpace(varData)) varData = "0";

            try
            {

                dData = Convert.ToDouble(varData);

            }
            catch (Exception ex)
            {
                VoiceAttackPlugin.LogOutput("WASM method failed.", "grey");
                return;
            }



            if (String.IsNullOrWhiteSpace(varData)) varData = "0";

            //------CALCULATOR CODE--------------------------------------------------------------

            string calcCode = varData + " (>L:" + varName + ")";


            _waSimConnection.executeCalculatorCode(calcCode, 0, out double fResult, out string sResult);

            Debug.WriteLine($"Calculator code '{calcCode}' returned: {fResult} and '{sResult}'", "<<");

            //----------------------------------------------------------------------------------

            //_waSimConnection.setVariable(new VariableRequest(varName), dData);

            //_waSimConnection.setLocalVariable(varName, dData);

            Debug.WriteLine("setLocalVariable() variable sent...");


        }

        public double GetLVarFSUIPC(string varName)
        {
            VoiceAttackPlugin.LogOutput("Reading Local Variable through FSUIPC...", "grey");

            double varResult;

            FSUIPCConnection.Open();

            varResult = FSUIPCConnection.ReadLVar(varName);

            FSUIPCConnection.Close();

            return varResult;



        }

        public void SetLVarFSUIPC(string varName, string varData = "0")
        {

            VoiceAttackPlugin.LogOutput("Writing Local Variable through FSUIPC...", "grey");

            if (String.IsNullOrWhiteSpace(varData)) varData = "0";

            try
            {


                double dData = double.Parse(varData);

                FSUIPCConnection.Open();

                FSUIPCConnection.WriteLVar(varName, dData);

                FSUIPCConnection.Close();



            }
            catch (Exception ex)
            {
                VoiceAttackPlugin.LogOutput("Failed to write data.", "grey");
                return;
            }


        }

        public double TriggerCalcCode(string calcCode = "0")
        {


            //------CALCULATOR CODE--------------------------------------------------------------

            VoiceAttackPlugin.LogOutput("Sending calculator code...", "grey");

            _waSimConnection.executeCalculatorCode(calcCode, CalcResultType.Double, out double fResult, out string sResult);

            VoiceAttackPlugin.LogOutput("Calculator code " + calcCode + " returned: " + fResult + " and " + sResult, "grey");
            //----------------------------------------------------------------------------------

            //_waSimConnection.setVariable(new VariableRequest(varName), dData);

            //_waSimConnection.setLocalVariable(varName, dData);



            return fResult;

        }
        /// <summary>
        /// Read variable through WASM
        /// </summary>
        /// <param name="varName"></param>
        /// <returns></returns>
        public double TriggerWASM(string varName)
        {


            //------CALCULATOR CODE--------------------------------------------------------------

            VoiceAttackPlugin.LogOutput("Triggering WASM read method...", "grey");

            string calcCode = "(L:" + varName + ")";

            if (varName == "A32NX_AUTOPILOT_HEADING_SELECTED")
            {

                readsSep = true;

            }

            _waSimConnection.executeCalculatorCode(calcCode, CalcResultType.Double, out double fResult, out string sResult);


            //-----------------------------------------------------------------------------------

            return fResult;

        }

        /// <summary>
        /// Send PMDG 737 variable through FSUIPC
        /// </summary>
        /// <param name="varName"></param>
        /// <param name="varData"></param>
        public void TriggerPMDG(string varName, string varData = "0")
        {
            VoiceAttackPlugin.LogOutput("Triggering PMDG write method...", "grey");

            if (String.IsNullOrWhiteSpace(varData)) varData = "0";

            try
            {


                int ivarData = int.Parse(varData);

                FSUIPCConnection.Open();

                var val = (int)Enum.Parse(typeof(PMDG_737_NGX_Control), varName);

                FSUIPCConnection.SendControlToFS(val, ivarData);

                FSUIPCConnection.Close();



            }
            catch (Exception ex)
            {
                VoiceAttackPlugin.LogOutput("Failed to send data.", "grey");
                return;
            }


        }

        public string TriggerReqPMDG(string varName)
        {
            VoiceAttackPlugin.LogOutput("Triggering PMDG read method...", "grey");

            try
            {

                string valueString = "NULL";

                string variableType = "Bool";

                int variableLength = 0;

                FSUIPCConnection.Open();

                int address = (int)Enum.Parse(typeof(PMDGConditionVariables.PMDGVarAddress), varName);

                int variableTypeC = (int)Enum.Parse(typeof(PMDGConditionVariables.PMDGVarTypes), varName);

                switch (variableTypeC)
                {
                    case 11:

                        variableType = "BYTE";
                        variableLength = 1;

                        break;

                    case 12:

                        variableType = "BYTEx2";
                        variableLength = 2;

                        break;

                    case 13:

                        variableType = "BYTEx3";
                        variableLength = 3;

                        break;

                    case 14:

                        variableType = "BYTEx4";
                        variableLength = 4;

                        break;

                    case 17:

                        variableType = "BYTEx7";
                        variableLength = 7;

                        break;

                    case 18:

                        variableType = "BYTEx8";
                        variableLength = 8;

                        break;

                    case 116:

                        variableType = "BYTEx16";
                        variableLength = 16;

                        break;

                    case 20:

                        variableType = "SHORT";
                        variableLength = 2;

                        break;

                    case 21:

                        variableType = "UINT";
                        variableLength = 4;

                        break;

                    case 22:

                        variableType = "INT";
                        variableLength = 4;

                        break;

                    case 41:

                        variableType = "FLT32";
                        variableLength = 4;

                        break;

                    case 42:

                        variableType = "FLT32x2";
                        variableLength = 8;

                        break;

                    case 56:

                        variableType = "CHARx6";
                        variableLength = 6;

                        break;

                    case 513:

                        variableType = "CHARx13";
                        variableLength = 13;

                        break;

                    case 601:

                        variableType = "DWORD";
                        variableLength = 4;

                        break;

                    case 603:

                        variableType = "DWORDx3";
                        variableLength = 12;

                        break;

                    case 611:

                        variableType = "WORD";
                        variableLength = 2;

                        break;

                    case 612:

                        variableType = "WORDx2";
                        variableLength = 4;

                        break;

                    case 7:

                        variableType = "STRING";
                        variableLength = 9;

                        break;

                }

                Offset myOffset = new Offset("", address, variableLength, false);

                FSUIPCConnection.Process();

                switch (variableType)
                {
                    case "BYTE":
                    case "BYTEx2":
                    case "BYTEx3":
                    case "BYTEx4":
                    case "BYTEx7":
                    case "BYTEx8":
                    case "BYTEx16":

                        byte[] valueByte2 = myOffset.GetValue<byte[]>();

                        valueString = BitConverter.ToString(valueByte2);

                        break;

                    case "SHORT":

                        short valueShort = myOffset.GetValue<short>();

                        valueString = Convert.ToString(valueShort);

                        break;

                    case "INT":

                        int valueInt = myOffset.GetValue<int>();

                        valueString = Convert.ToString(valueInt);

                        break;

                    case "UINT":
                    case "DWORD":
                    case "DWORDx3":

                        uint valueUint = myOffset.GetValue<uint>();

                        valueString = Convert.ToString(valueUint);

                        break;

                    case "FLT32":
                    case "FLT32x2":

                        float valueFloat = myOffset.GetValue<float>();

                        valueString = Convert.ToString(valueFloat);

                        break;

                    case "CHARx6":
                    case "CHARx13":
                    case "STRING":

                        valueString = myOffset.GetValue<string>();

                        break;

                    case "WORD":
                    case "WORDx2":

                        ushort valueUshort = myOffset.GetValue<ushort>();

                        valueString = Convert.ToString(valueUshort);

                        break;


                }


                myOffset.Disconnect();

                FSUIPCConnection.Close();

                return valueString;

            }
            catch (Exception ex)
            {
                VoiceAttackPlugin.LogOutput("Failed to receive data.", "grey");
                return "NULL";
            }


        }
        public void TriggerComKey(string varName, string varData = "0")
        {

            VoiceAttackPlugin.LogOutput("Triggering special Key Event method.", "grey");

            if (varName == "XPNDR_SET")
            {
                Decimal d;

                UInt32 varDataBCD;

                Decimal.TryParse(varData, out d);

                varDataBCD = Utils.Dec2Bcd(Decimal.ToUInt32(d));

                varData = varDataBCD.ToString();

            }


            if (varName == "ADF_COMPLETE_SET" | varName == "ADF2_COMPLETE_SET" | varName == "COM_RADIO_SET" | varName == "COM2_RADIO_SET" | varName == "COM_STBY_RADIO_SET" | varName == "COM2_STBY_RADIO_SET" | varName == "NAV1_RADIO_SET" | varName == "NAV2_RADIO_SET" | varName == "NAV1_STBY_SET" | varName == "NAV2_STBY_SET")
            {
                Decimal d;

                UInt32 varDataBCD;

                Decimal.TryParse(varData, out d);

                varDataBCD = Utils.Dec2Bcd(Decimal.ToUInt32(d * 100));

                varData = varDataBCD.ToString();

            }

            string calcCode = varData + " (>K:" + varName + ")";



            _waSimConnection.executeCalculatorCode(calcCode, 0, out double fResult, out string sResult);

            VoiceAttackPlugin.LogOutput("Calculator code " + calcCode + " returned: " + fResult + " and " + sResult, "grey");

            

            //----------------------------------------------------------------------------------

            //_waSimConnection.setVariable(new VariableRequest(varName), dData);

            //_waSimConnection.setLocalVariable(varName, dData);



        }

        public void TriggerKey(string varName, string varData = "0")
        {
            VoiceAttackPlugin.LogOutput("Triggering Key Event method.", "grey");

            if (String.IsNullOrWhiteSpace(varData)) varData = "0";

            try
            {


                int ivarData = int.Parse(varData);

                FSUIPCConnection.Open();

                var val = (int)Enum.Parse(typeof(FsControl), varName);

                FSUIPCConnection.SendControlToFS(val, ivarData);

                FSUIPCConnection.Close();



            }
            catch (Exception ex)
            {
                VoiceAttackPlugin.LogOutput("Failed to send event data.", "grey");
                return;
            }


        }

        public void TriggerKeySimconnect(FsControlList varName, string varData = "0")
        {

            VoiceAttackPlugin.LogOutput("Triggering Key Event method.", "grey");

            UInt32 kData;
            Decimal d;

            if (String.IsNullOrWhiteSpace(varData)) varData = "0";

            try
            {                                

                        Byte[] Bytes = BitConverter.GetBytes(Convert.ToInt32(varData));
                        kData = BitConverter.ToUInt32(Bytes, 0);

                    
                
            }
            catch (Exception ex)
            {
                VoiceAttackPlugin.LogOutput("Failed to send event data.", "grey");
                return;
            }

            _simConnection.TransmitClientEvent(SimConnect.SIMCONNECT_OBJECT_ID_USER, varName, kData, NOTIFICATION_GROUPS.DEFAULT, SIMCONNECT_EVENT_FLAG.GROUPID_IS_PRIORITY);

            VoiceAttackPlugin.LogOutput("Data sent.", "grey");

        }

        public double TriggerReqSimVar(string varName)
        {
            VoiceAttackPlugin.LogOutput("Triggering SimVar read method...", "grey");

            string unit = "NULL";

            char last = varName[varName.Length - 1];

            bool isIndex = Char.IsDigit(last);

            SimVarFunc simVarFunc = new SimVarFunc();

            if (isIndex)
            {
                varName = varName.Remove(varName.Length - 1);

                char blast = varName[varName.Length - 1];

                bool isbIndex = Char.IsDigit(blast);

                if (isbIndex)
                {

                    varName = varName.Remove(varName.Length - 1);

                    char clast = varName[varName.Length - 1];

                    bool iscIndex = Char.IsDigit(clast);

                    if (iscIndex)
                    {
                        varName = varName.Remove(varName.Length - 2);

                        varName = varName.ToUpper();

                        int variableUnitID = (int)Enum.Parse(typeof(SimVarList), varName);

                        if (variableUnitID == 13 | variableUnitID == 6)
                        {
                            readsSep = true;

                        }

                        varName = varName.Replace("_", " ");

                        unit = simVarFunc.GiveSimVarUnit(variableUnitID);

                        string calcCode = "(A:" + varName + ":" + clast + blast + last + ", " + unit + ")";

                        _waSimConnection.executeCalculatorCode(calcCode, CalcResultType.Double, out double fResult, out string sResult);

                        Debug.WriteLine("Executed calculator code with 3 numbers: " + calcCode);

                        VoiceAttackPlugin.LogOutput("Executed calculator code with 3 index digits: " + calcCode, "grey");

                        if (varName == "TRANSPONDER CODE")
                        {

                            fResult = fResult * 1000000;

                        }

                        return fResult;
                    }
                    else
                    {
                        varName = varName.Remove(varName.Length - 1);

                        varName = varName.ToUpper();

                        int variableUnitID = (int)Enum.Parse(typeof(SimVarList), varName);

                        if (variableUnitID == 13 | variableUnitID == 6)
                        {
                            readsSep = true;
                        }

                        varName = varName.Replace("_", " ");

                        unit = simVarFunc.GiveSimVarUnit(variableUnitID);

                        string calcCode = "(A:" + varName + ":" + blast + last + ", " + unit + ")";

                        _waSimConnection.executeCalculatorCode(calcCode, CalcResultType.Double, out double fResult, out string sResult);

                        VoiceAttackPlugin.LogOutput("Executed calculator code with 2 index digits: " + calcCode, "grey");

                        if (varName == "TRANSPONDER CODE")
                        {

                            fResult = fResult * 1000000;

                        }

                        return fResult;
                    }


                }
                else
                {

                    varName = varName.Remove(varName.Length - 1);

                    varName = varName.ToUpper();

                    int variableUnitID = (int)Enum.Parse(typeof(SimVarList), varName);

                    if (variableUnitID == 13 | variableUnitID == 6)
                    {
                        readsSep = true;
                    }

                    varName = varName.Replace("_", " ");

                    unit = simVarFunc.GiveSimVarUnit(variableUnitID);

                    string calcCode = "(A:" + varName + ":" + last + ", " + unit + ")";

                    _waSimConnection.executeCalculatorCode(calcCode, CalcResultType.Double, out double fResult, out string sResult);

                    VoiceAttackPlugin.LogOutput("Executed calculator code with 1 index digit: " + calcCode, "grey");

                    if (varName == "TRANSPONDER CODE")
                    {

                        fResult = fResult * 1000000;

                    }

                    return fResult;

                }

            }
            else
            {

                varName = varName.ToUpper();

                int variableUnitID = (int)Enum.Parse(typeof(SimVarList), varName);

                if (variableUnitID == 13 | variableUnitID == 6)
                {
                    readsSep = true;
                }

                varName = varName.Replace("_", " ");

                unit = simVarFunc.GiveSimVarUnit(variableUnitID);

                string calcCode = "(A:" + varName + ", " + unit + ")";

                _waSimConnection.executeCalculatorCode(calcCode, CalcResultType.Double, out double fResult, out string sResult);


                VoiceAttackPlugin.LogOutput("Executed calculator code: " + calcCode, "grey");

                if (varName == "TRANSPONDER CODE")
                {

                    fResult = fResult * 1000000;

                }

                return fResult;


            }


        }
        public bool SBFetch()
        {
            bool fetchRes;



            VoiceAttackPlugin.LogOutput("SimBrief fetch in started...", "grey");

            SimBriefFetch simbriefFetch = new SimBriefFetch();

            fetchRes = simbriefFetch.Load();


            VoiceAttackPlugin.LogOutput("SimBrief data received: ", "blue");

            TriggerCalcCode(Utils.sbCostIndex + " (>L:SimBrief_CostIndex)");
            VoiceAttackPlugin.LogOutput("CostIndex: " + Utils.sbCostIndex, "blue");

            TriggerCalcCode(Utils.sbOriginElevation + " (>L:SimBrief_OriginElevation)");
            VoiceAttackPlugin.LogOutput("OriginElevation: " + Utils.sbOriginElevation, "blue");

            TriggerCalcCode(Utils.sbOriginTransAlt + " (>L:SimBrief_OriginTransAlt)");
            VoiceAttackPlugin.LogOutput("OriginTransAlt: " + Utils.sbOriginTransAlt, "blue");

            TriggerCalcCode(Utils.sbOriginTransLevel + " (>L:SimBrief_OriginTransLevel)");
            VoiceAttackPlugin.LogOutput("OriginTransLevel: " + Utils.sbOriginTransLevel, "blue");

            TriggerCalcCode(Utils.sbOriginWindDir + " (>L:SimBrief_OriginWindDir)");
            VoiceAttackPlugin.LogOutput("OriginWindDir: " + Utils.sbOriginWindDir, "blue");

            TriggerCalcCode(Utils.sbOriginWindSpd + " (>L:SimBrief_OriginWindSpd)");
            VoiceAttackPlugin.LogOutput("OriginWindSpd: " + Utils.sbOriginWindSpd, "blue");

            TriggerCalcCode(Utils.sbOriginQNH + " (>L:SimBrief_OriginQNH)");
            VoiceAttackPlugin.LogOutput("OriginQNH: " + Utils.sbOriginQNH, "blue");

            TriggerCalcCode(Utils.sbOriginBaro + " (>L:SimBrief_OriginBaro)");
            VoiceAttackPlugin.LogOutput("OriginBaro: " + Utils.sbOriginBaro, "blue");

            TriggerCalcCode(Utils.sbAltnElevation + " (>L:SimBrief_AltnElevation)");
            VoiceAttackPlugin.LogOutput("AltnElevation: " + Utils.sbAltnElevation, "blue");

            TriggerCalcCode(Utils.sbAltnTransAlt + " (>L:SimBrief_AltnTransAlt)");
            VoiceAttackPlugin.LogOutput("AltnTransAlt: " + Utils.sbAltnTransAlt, "blue");

            TriggerCalcCode(Utils.sbAltnTransLevel + " (>L:SimBrief_AltnTransLevel)");
            VoiceAttackPlugin.LogOutput("AltnTransLevel: " + Utils.sbAltnTransLevel, "blue");

            TriggerCalcCode(Utils.sbAltnWindDir + " (>L:SimBrief_AltnWindDir)");
            VoiceAttackPlugin.LogOutput("AltnWindDir: " + Utils.sbAltnWindDir, "blue");

            TriggerCalcCode(Utils.sbAltnWindSpd + " (>L:SimBrief_AltnWindSpd)");
            VoiceAttackPlugin.LogOutput("AltnWindSpd: " + Utils.sbAltnWindSpd, "blue");

            TriggerCalcCode(Utils.sbAltnQNH + " (>L:SimBrief_AltnQNH)");
            VoiceAttackPlugin.LogOutput("AltnQNH: " + Utils.sbAltnQNH, "blue");

            TriggerCalcCode(Utils.sbAltnBaro + " (>L:SimBrief_AltnBaro)");
            VoiceAttackPlugin.LogOutput("AltnBaro: " + Utils.sbAltnBaro, "blue");

            TriggerCalcCode(Utils.sbDestElevation + " (>L:SimBrief_DestElevation)");
            VoiceAttackPlugin.LogOutput("DestElevation: " + Utils.sbDestElevation, "blue");

            TriggerCalcCode(Utils.sbDestTransAlt + " (>L:SimBrief_DestTransAlt)");
            VoiceAttackPlugin.LogOutput("DestTransAlt: " + Utils.sbDestTransAlt, "blue");

            TriggerCalcCode(Utils.sbDestTransLevel + " (>L:SimBrief_DestTransLevel)");
            VoiceAttackPlugin.LogOutput("DestTransLevel: " + Utils.sbDestTransLevel, "blue");

            TriggerCalcCode(Utils.sbDestWindDir + " (>L:SimBrief_DestWindDir)");
            VoiceAttackPlugin.LogOutput("DestWindDir: " + Utils.sbDestWindDir, "blue");

            TriggerCalcCode(Utils.sbDestWindSpd + " (>L:SimBrief_DestWindSpd)");
            VoiceAttackPlugin.LogOutput("DestWindSpd: " + Utils.sbDestWindSpd, "blue");

            TriggerCalcCode(Utils.sbDestQNH + " (>L:SimBrief_DestQNH)");
            VoiceAttackPlugin.LogOutput("DestQNH: " + Utils.sbDestQNH, "blue");

            TriggerCalcCode(Utils.sbDestBaro + " (>L:SimBrief_DestBaro)");
            VoiceAttackPlugin.LogOutput("DestBaro: " + Utils.sbDestBaro, "blue");

            TriggerCalcCode(Utils.sbFinRes + " (>L:SimBrief_FinRes)");
            VoiceAttackPlugin.LogOutput("FinRes: " + Utils.sbFinRes, "blue");

            TriggerCalcCode(Utils.sbAltnFuel + " (>L:SimBrief_AltnFuel)");
            VoiceAttackPlugin.LogOutput("AltnFuel: " + Utils.sbAltnFuel, "blue");

            TriggerCalcCode(Utils.sbFinresPAltn + " (>L:SimBrief_FinresPAltn)");
            VoiceAttackPlugin.LogOutput("FinresPAltn: " + Utils.sbFinresPAltn, "blue");

            TriggerCalcCode(Utils.sbFuel + " (>L:SimBrief_Fuel)");
            VoiceAttackPlugin.LogOutput("Fuel: " + Utils.sbFuel, "blue");

            TriggerCalcCode(Utils.sbPassenger + " (>L:SimBrief_Passenger)");
            VoiceAttackPlugin.LogOutput("Passenger: " + Utils.sbPassenger, "blue");

            TriggerCalcCode(Utils.sbBags + " (>L:SimBrief_Bags)");
            VoiceAttackPlugin.LogOutput("Bags: " + Utils.sbBags, "blue");

            TriggerCalcCode(Utils.sbWeightPax + " (>L:SimBrief_WeightPax)");
            VoiceAttackPlugin.LogOutput("WeightPax: " + Utils.sbWeightPax, "blue");

            TriggerCalcCode(Utils.sbWeightCargo + " (>L:SimBrief_WeightCargo)");
            VoiceAttackPlugin.LogOutput("WeightCargo: " + Utils.sbWeightCargo, "blue");

            TriggerCalcCode(Utils.sbPayload + " (>L:SimBrief_Payload)");
            VoiceAttackPlugin.LogOutput("Payload: " + Utils.sbPayload, "blue");

            TriggerCalcCode(Utils.sbZFW + " (>L:SimBrief_sbZFW)");
            VoiceAttackPlugin.LogOutput("sbZFW: " + Utils.sbZFW, "blue");

            TriggerCalcCode(Utils.sbTOW + " (>L:SimBrief_sbTOW)");
            VoiceAttackPlugin.LogOutput("sbTOW: " + Utils.sbTOW, "blue");


            VoiceAttackPlugin.LogOutput("SimBrief fetch finished.", "grey");

            return fetchRes;

        }
        #endregion

        #region Event Handlers

        /// <summary>
        /// Initializes all the event handles for the agent
        /// </summary>

        private void initWASMHandlers()
        {

            _waSimConnection.OnClientEvent += ClientStatusHandler;
            _waSimConnection.OnLogRecordReceived += LogHandler;
            _waSimConnection.OnDataReceived += DataSubscriptionHandler;
            _waSimConnection.setLogLevel(LogLevel.Trace, LogFacility.File, LogSource.Client);
            _waSimConnection.setLogLevel(LogLevel.Trace, LogFacility.Remote, LogSource.Client);
            _waSimConnection.setLogLevel(LogLevel.Trace, LogFacility.File, LogSource.Server);
            _waSimConnection.setLogLevel(LogLevel.Trace, LogFacility.Remote, LogSource.Server);

        }

        /// <summary>
        /// Catches any exceptions that are encountered by SimConnect
        /// </summary>
        private void simconnect_OnRecvException(SimConnect sender, SIMCONNECT_RECV_EXCEPTION data)
        {
            SIMCONNECT_EXCEPTION ex = (SIMCONNECT_EXCEPTION)data.dwException;


            // A common exception will be unrecognized data definitions or events

            // Info on the "data" returned with the exception:
            // - dwException enum type of SIMCONNECT_EXCEPTION
            // - "UNKNOWN_SENDID" not sure
            // - dwSendID  # see SimConnect_GetLastSentPacketID
            // - "UNKNOWN_INDEX" not sure
            // - dwIndex # index of parameter that was source of error
        }

        /// <summary>
        /// Fires when the connection to the sim is successfully made
        /// </summary>
        private void simconnect_OnRecvOpen(SimConnect sender, SIMCONNECT_RECV_OPEN data)
        {


            Connected = true;

        }

        /// <summary>
        /// Fires when the sim is closed/exited.  
        /// </summary>
        private void simconnect_OnRecvQuit(SimConnect sender, SIMCONNECT_RECV data)
        {

            Disconnect();
        }

        /// <summary>
        /// Fires when new data is recieved from the sim
        /// </summary>
        private void simconnect_OnRecvSimobjectDataBytype(SimConnect sender, SIMCONNECT_RECV_SIMOBJECT_DATA_BYTYPE data)

        {



        }



        /// <summary>
        /// Handles the polling for new messages from the sim
        /// </summary>
        private void SimConnect_MessageReceiveThreadHandler()
        {
            while (true)
            {
                // interval is in milliseconds (may want to increase this)
                _simConnectEventHandle.WaitOne(1);

                try
                {
                    // ask the sim if it has any messages (data responses)
                    _simConnection?.ReceiveMessage();
                }
                catch (Exception ex)
                {
                    VoiceAttackPlugin.LogOutput("MessageReceiveThreadHandler failed.", "grey");
                }
            }
        }

        /// <summary>
        /// A manual way to check for messages from the sim
        /// </summary>
        public void CheckForMessage()
        {

            // ask the sim if it has any messages (data responses)
            _simConnection.ReceiveMessage();
            VoiceAttackPlugin.LogOutput("Checked for messages (manually).", "grey");
        }

        static void ClientStatusHandler(ClientEvent ev)
        {

        }

        static void LogHandler(LogRecord lr, LogSource src)
        {

        }
        #endregion

        private void initEventHandlers()
        {
            try
            {
                _simConnection.OnRecvOpen += new SimConnect.RecvOpenEventHandler(simconnect_OnRecvOpen);
                _simConnection.OnRecvQuit += new SimConnect.RecvQuitEventHandler(simconnect_OnRecvQuit);
                _simConnection.OnRecvException += new SimConnect.RecvExceptionEventHandler(simconnect_OnRecvException);
                _simConnection.OnRecvEvent += new SimConnect.RecvEventEventHandler(simconnect_OnRecvEvent);
                _simConnection.OnRecvSimobjectDataBytype += new SimConnect.RecvSimobjectDataBytypeEventHandler(simconnect_OnRecvSimobjectDataBytype);

                _simConnection.SubscribeToSystemEvent(FsControlList.SIMSTART, "SimStart");
                _simConnection.SubscribeToSystemEvent(FsControlList.SIMSTOP, "SimStop");

                // I can't recall why this is needed
                _simConnection.SetNotificationGroupPriority(NOTIFICATION_GROUPS.DEFAULT, SimConnect.SIMCONNECT_GROUP_PRIORITY_HIGHEST);

            }
            catch (Exception ex)
            {
                VoiceAttackPlugin.LogOutput("Failed to initiate event handlers", "grey");
            }
        }

        

        private void simconnect_OnRecvEvent(SimConnect sender, SIMCONNECT_RECV_EVENT recEvent)
        {

            // I'm still try to figure out when these get triggered.  
            // They often fire rapidly when exiting and entering flights.            

            switch ((FsControlList)recEvent.uEventID)
            {
                case FsControlList.SIMSTART:

                    Debug.WriteLine("Sim running");
                    break;

                case FsControlList.SIMSTOP:

                    Debug.WriteLine("Sim stopped");
                    break;

                case FsControlList.PAUSE:

                    Debug.WriteLine("Sim paused");
                    break;

            }
        }
        private enum Requests : uint
        {
            REQUEST_ID_1_FLOAT,
            REQUEST_ID_2_STR
        }

        static void DataSubscriptionHandler(DataRequestRecord dr)
        {

            // Convert the received data into a value using DataRequestRecord's tryConvert() methods.
            // This could be more efficient in a "real" application, but it's good enough for our tests with only 2 value types.
            if (dr.tryConvert(out float fVal)) ;

            else if (dr.tryConvert(out string sVal))
            {

            }
            else;


        }

    }

}
