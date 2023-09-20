//=================================================================================================================
// PROJECT: MSFS VAPlugin
// PURPOSE: This class does the interfacing to SimConnect and WASM for Microsoft Flight Simulator 2020
// AUTHOR: William Riker
//================================================================================================================= 
using System;
using System.Collections.Generic;
using System.Threading;
using System.Diagnostics;
using System.Timers;
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
        public bool isLastMessageReceived = false;


        public static uint currentDefinition = 1;
        public static uint currentRequest = 1;

        public static System.Timers.Timer checkMessagesTimer = new System.Timers.Timer(500);

        private static IntPtr m_hWnd = new IntPtr(0);

        public static string requestType = "NULL";
        public static string airportIcao = "NULL";
        public static string runwayDesignator = "NULL";

        public double simvardata;

        SimVarDouble _simVarD = new SimVarDouble();
        SimVarString _simVarS = new SimVarString();
        SimVarInt _simVarI = new SimVarInt();
        SimVarBool _simVarB = new SimVarBool();
        SimVarFloat _simVarF = new SimVarFloat();

        public SimVarDouble GetSimVarD
        {
            get => _simVarD;
            private set
            {
                _simVarD = value;
            }
        }

        public SimVarString GetSimVarS
        {
            get => _simVarS;
            private set
            {
                _simVarS = value;
            }
        }

        public SimVarFloat GetSimVarF
        {
            get => _simVarF;
            private set
            {
                _simVarF = value;
            }
        }

        public SimVarInt GetSimVarI
        {
            get => _simVarI;
            private set
            {
                _simVarI = value;
            }
        }

        public SimVarBool GetSimVarB
        {
            get => _simVarB;
            private set
            {
                _simVarB = value;
            }
        }

        public enum WasmModuleStatus
        {
            Unknown, NotFound, Found, Connected
        }

        public bool WasmAvailable => WasmStatus != WasmModuleStatus.NotFound;
        public WasmModuleStatus WasmStatus { get; private set; } = WasmModuleStatus.Unknown;


        Dictionary<RequestTypes, bool> requestPending = new Dictionary<RequestTypes, bool>();
        PlaneState _planeState = new PlaneState();
        Dictionary<FacilityRequestTypes, bool> requestPendingFac = new Dictionary<FacilityRequestTypes, bool>();


        // used for polling for messages from the sim
        EventWaitHandle _simConnectEventHandle = new EventWaitHandle(false, EventResetMode.AutoReset);
        Thread _simConnectReceiveThread = null;

        /// <summary>
        /// Returns the last known state of the plane
        /// </summary>


        /// <summary>
        /// Checks if a Request is still waiting for a response from the sim
        /// </summary>
        public bool RequestPending(RequestTypes requestType)
        {
            return requestPending[requestType];
        }

        public bool RequestPendingFac(FacilityRequestTypes requestType)
        {
            return requestPendingFac[requestType];
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
                    _simConnection = new Microsoft.FlightSimulator.SimConnect.SimConnect("FVCplugin", IntPtr.Zero, WM_USER_SIMCONNECT, null, 0);
                    initEventHandlers();
                    initEvents();
                    Connected = true;
                }
                catch (Exception e)
                {
                    VoiceAttackPlugin.LogErrorOutput("SimConnect connection failed. " + e.Message, "red");
                    Connected = false;
                }

            }

        }

        public void Connectfacility()
        {


                try
                {
                    VoiceAttackPlugin.LogOutput("Trying to connect to facility...", "grey");

                    _simConnection = new SimConnect("FVCplugin", m_hWnd, WM_USER_SIMCONNECT, null, 0);

                    /// Listen to connect and quit msgs
                    _simConnection.OnRecvOpen += new SimConnect.RecvOpenEventHandler(simconnect_Facility_OnRecvOpen);



                    _simConnection.OnRecvQuit += new SimConnect.RecvQuitEventHandler(simconnect_Facility_OnRecvQuit);
                    _simConnection.OnRecvException += new SimConnect.RecvExceptionEventHandler(simconnect_OnRecvException);
                    _simConnection.OnRecvEvent += new SimConnect.RecvEventEventHandler(simconnect_OnRecvEvent);
                    _simConnection.OnRecvSimobjectDataBytype += new SimConnect.RecvSimobjectDataBytypeEventHandler(simconnect_OnRecvSimobjectDataBytype);

                    _simConnection.OnRecvFacilityData += new SimConnect.RecvFacilityDataEventHandler(simconnect_Facility_OnRecvFacilityData);
                    _simConnection.OnRecvFacilityDataEnd += new SimConnect.RecvFacilityDataEndEventHandler(simconnect_Facility_OnRecvFacilityDataEnd);

                    _simConnection.SetNotificationGroupPriority(NOTIFICATION_GROUPS.DEFAULT, SimConnect.SIMCONNECT_GROUP_PRIORITY_HIGHEST);

                    VoiceAttackPlugin.LogOutput("Facility connection established.", "grey");


                    Connected = true;
                }
                catch (Exception e)
                {
                    VoiceAttackPlugin.LogErrorOutput("SimConnect facility connection failed. " + e.Message, "red");
                    Connected = false;
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
                catch (Exception e)
                {
                    VoiceAttackPlugin.LogErrorOutput("WASM connection failed. " + e.Message, "red");
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
                catch (Exception e)
                {
                    VoiceAttackPlugin.LogErrorOutput("WASM connection failed. " + e.Message, "red");
                }
            }



        }

        public void WASMConnect3()
        {
            try
            {
                _waSimConnection = new WASimCommander.CLI.Client.WASimClient(0xC57E57E9);
                _waSimConnection.connectSimulator();
                _waSimConnection.connectServer();

                for (int i = 0; i < 10 && _waSimConnection.isConnected()==false; i++)
                {
                    VoiceAttackPlugin.LogOutput("WASM connecting...", "grey");
                    _waSimConnection.connectSimulator();
                    _waSimConnection.connectServer();
                }
                WAServerConnected = _waSimConnection.isConnected();
                if (WAServerConnected == false)
                {
                    VoiceAttackPlugin.LogOutput("WASM not able to connect.", "grey");
                    MSFS.Utils.errcon = true;
                }
                else
                {
                    MSFS.Utils.errcon = false;
                }

            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogErrorOutput("WASM connection failed. " + e.Message, "red");
            }            
        }

        public void WASMConnect2()
        {

            if (_waSimConnection == null)
            {

                VoiceAttackPlugin.LogErrorOutput("WASM client not initialized.", "red");

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

                VoiceAttackPlugin.LogErrorOutput("WASM Client could not connect to SimConnect for unknown reason.", "grey");

            if (hr != HR.OK)
            {
                WasmStatus = WasmModuleStatus.NotFound;
                VoiceAttackPlugin.LogErrorOutput("WASM Server not found or couldn't connect.", "grey");
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
                catch (Exception e)
                {
                    VoiceAttackPlugin.LogErrorOutput("Unable to connect to WASM. " + e.Message, "red");
                }
            }



        }



        /// <summary>
        /// Starts a background thread to periodically check for messages back from the sim
        /// </summary>
        public void EnableMessagePolling()
        {

            try
            {
                _simConnectReceiveThread = new Thread(new ThreadStart(SimConnect_MessageReceiveThreadHandler));
                _simConnectReceiveThread.IsBackground = true;
                _simConnectReceiveThread.Start();

            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogErrorOutput("Enable Message Polling failed. " + e.Message, "red");
            }
        }

        /// <summary>
        /// Turns of background thread used message polling with the sim
        /// </summary>
        public void DisableMessagePolling()
        {

            if (_simConnectReceiveThread != null)
            {
                try
                {
                    _simConnectReceiveThread.Abort();
                    _simConnectReceiveThread.Join();
                    _simConnectReceiveThread = null;

                }
                catch (Exception e)
                {
                    VoiceAttackPlugin.LogErrorOutput("Disable Message Polling failed. " + e.Message, "red");
                }

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

                VoiceAttackPlugin.LogOutput("Connection to sim closed successfully.", "grey");
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogErrorOutput("Failed to Disconnect and clean up. " + e.Message, "red");
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
                try
                {
                    _simConnection.MapClientEventToSimEvent(item, keyName);
                    _simConnection.AddClientEventToNotificationGroup(NOTIFICATION_GROUPS.DEFAULT, item, false);

                }
                catch (Exception e)
                {
                    VoiceAttackPlugin.LogErrorOutput("Initialize events failed. " + e.Message, "red");
                }

            }

        }


        public void WASMDisconnect()
        {

            if (!WAServerConnected) return;
            
            try
            {

                _waSimConnection.disconnectServer();
                _waSimConnection.disconnectSimulator();

                for (int i = 0; i < 10 && _waSimConnection.isConnected() == true; i++)
                {
                    VoiceAttackPlugin.LogOutput("WASM disconnecting...", "grey");
                }

                WAServerConnected = _waSimConnection.isConnected();

                if (WAServerConnected == true)
                {
                    VoiceAttackPlugin.LogOutput("WASM not able to disconnect.", "grey");
                }
                else
                {
                    VoiceAttackPlugin.LogOutput("Connection to WASM successfully closed.", "grey");
                }

                // delete the client

                _waSimConnection.Dispose();
                
                Thread.Sleep(200);
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogErrorOutput("Failed to Disconnect WASM. " + e.Message, "red");
            }
            finally
            {

                _waSimConnection = null;

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
            try
            {
                //colors don't seem to be supported in msfs
                _simConnection.Text(SIMCONNECT_TEXT_TYPE.PRINT_WHITE, duration, null, text);
            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogErrorOutput("Set Text failed. " + e.Message, "red");
            }

        }



        /// <summary>
        /// Set variable through WASM
        /// </summary>
        /// <param name="varName"></param>
        /// <param name="varData"></param>
        public void TriggerWASM(string varName, string varData = "0")
        {
            VoiceAttackPlugin.LogOutput("WASM L variable write method triggered...", "grey");

            double dData;

            if (String.IsNullOrWhiteSpace(varData)) varData = "0";
            string calcCode = varData + " (>L:" + varName + ")";
            dData = Convert.ToDouble(varData);

            try
            {

                _waSimConnection.executeCalculatorCode(calcCode, 0, out double fResult, out string sResult);
                VoiceAttackPlugin.LogOutput($"Calculator code '{calcCode}' returned: {fResult} and '{sResult}'", "grey");

            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogErrorOutput("WASM L variable set failed. " + e.Message, "red");
                return;
            }


        }

        public double GetLVarFSUIPC(string varName)
        {
            VoiceAttackPlugin.LogOutput("Reading Local Variable through FSUIPC...", "grey");

            double varResult;
            try
            {

                FSUIPCConnection.Open();

                try
                {

                    varResult = FSUIPCConnection.ReadLVar(varName);

                    if (varResult == 0)
                    {

                        varResult = FSUIPCConnection.ReadLVar(varName);

                        if (varResult == 0)
                        {

                            varResult = FSUIPCConnection.ReadLVar(varName);

                        }
                    }

                    try
                    {

                        FSUIPCConnection.Close();

                    }
                    catch (Exception e)
                    {
                        VoiceAttackPlugin.LogErrorOutput("FSUIPC couldn't disconnect. " + e.Message, "red");
                    }


                    return varResult;

                }
                catch (Exception e)
                {
                    VoiceAttackPlugin.LogErrorOutput("FSUIPC couldn't read L Var. " + e.Message, "red");
                    return 0;
                }

            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogErrorOutput("FSUIPC couldn't connect. " + e.Message, "red");

                FSUIPCConnection.Close();

                return 0;
            }               
                                  

        }

        public void SetLVarFSUIPC(string varName, string varData = "0")
        {

            VoiceAttackPlugin.LogOutput("Writing Local Variable through FSUIPC...", "grey");

            if (String.IsNullOrWhiteSpace(varData)) varData = "0";


            double dData = double.Parse(varData);

            try
            {

                FSUIPCConnection.Open();

            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogErrorOutput("FSUIPC couldn't connect. " + e.Message, "red");
            }
            try
            {

                FSUIPCConnection.WriteLVar(varName, dData);


            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogErrorOutput("FSUIPC couldn't write L Var. " + e.Message, "red");

            }


            try
            {

                FSUIPCConnection.Close();

            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogErrorOutput("FSUIPC couldn't disconnect. " + e.Message, "red");
            }



        }

        public double TriggerCalcCode(string calcCode = "0")
        {


            VoiceAttackPlugin.LogOutput("Calculator code method executed...", "grey");

            try
            {

                _waSimConnection.executeCalculatorCode(calcCode, CalcResultType.Double, out double fResult, out string sResult);
                VoiceAttackPlugin.LogOutput("Calculator code " + calcCode + " returned: " + fResult + " and " + sResult, "grey");
                return fResult;

            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogErrorOutput("Calculator code failed. " + e.Message, "red");
                return 0;
            }



        }
        /// <summary>
        /// Read variable through WASM
        /// </summary>
        /// <param name="varName"></param>
        /// <returns></returns>
        public double TriggerWASM(string varName)
        {


            VoiceAttackPlugin.LogOutput("WASM L variable read method triggered...", "grey");

            string calcCode = "(L:" + varName + ")";

            if (varName == "A32NX_AUTOPILOT_HEADING_SELECTED")
            {

                readsSep = true;

            }
            try
            {

                _waSimConnection.executeCalculatorCode(calcCode, CalcResultType.Double, out double fResult, out string sResult);
                VoiceAttackPlugin.LogOutput("Calculator code " + calcCode + " returned: " + fResult + " and " + sResult, "grey");
                return fResult;

            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogErrorOutput("Calculator code failed. " + e.Message, "red");
                return 0;
            }




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
            catch (Exception e)
            {
                VoiceAttackPlugin.LogErrorOutput("FSUIPC PMDG method failed. " + e.Message, "red");
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
            catch (Exception e)
            {
                VoiceAttackPlugin.LogOutput("FSUIPC PMDG failed to receive data. " + e.Message, "red");
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


            try
            {

                _waSimConnection.executeCalculatorCode(calcCode, 0, out double fResult, out string sResult);
                VoiceAttackPlugin.LogOutput("Calculator code " + calcCode + " returned: " + fResult + " and " + sResult, "grey");


            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogErrorOutput("Calculator code failed. " + e.Message, "red");

            }



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
            catch (Exception e)
            {
                VoiceAttackPlugin.LogErrorOutput("Failed to send key event data through FSUIPC. " + e.Message, "red");

            }


        }


        public void FacilityRequest()
        {

            VoiceAttackPlugin.LogOutput("Connecting Simconnect...", "grey");

            try
            {
                _simConnection = new SimConnect("FVCplugin", IntPtr.Zero, WM_USER_SIMCONNECT, null, 0);

                /// Listen to connect and quit msgs
                _simConnection.OnRecvOpen += simconnect_OnRecvOpen;
                _simConnection.OnRecvQuit += simconnect_OnRecvQuit;

                _simConnection.OnRecvFacilityData += simconnect_Facility_OnRecvFacilityData;
                _simConnection.OnRecvFacilityDataEnd += simconnect_Facility_OnRecvFacilityDataEnd;

                VoiceAttackPlugin.LogOutput("Connection established", "grey");

                
                while (isLastMessageReceived == false)
                {
                    System.Threading.Thread.Sleep(100);

                    _simConnection.ReceiveMessage();

                }
            }
            catch (Exception ex)
            {

                VoiceAttackPlugin.LogOutput("" + ex, "grey");
            }

            
        }

        public void FacilityWrappingUp()
        {


            

            VoiceAttackPlugin.LogOutput("Creating Local Variables...", "grey");

            float util;

            string pdesignator;

            string sdesignator;

            string adesignation;


            string bdesignation;


            string lvardesignation;

            try
            {


                switch (requestType)
                {

                    case "AIRPORT":

                        try
                        {

                        }
                        catch (Exception ex)
                        {
                            VoiceAttackPlugin.LogErrorOutput("Wrapping Up problem with AIRPORT type. " + ex, "red");
                        }


                        break;

                    case "RUNWAY":

                        pdesignator = "NULL";
                        sdesignator = "NULL";

                        bool match = false;


                        try
                        {
                            lvardesignation = VoiceAttackPlugin.GetText("MSFS.LvarRunwayDes");
                            VoiceAttackPlugin.LogOutput("LVar designator: " + lvardesignation, "grey");

                            VoiceAttackPlugin.LogOutput("Runway requested: " + Utils.reqRunway + ".", "grey");

                            VoiceAttackPlugin.SetText("MSFS." + lvardesignation + "_LOC_FREQ", "");
                            VoiceAttackPlugin.SetText("MSFS." + lvardesignation + "_LOC_HEAD", "");
                            VoiceAttackPlugin.SetText("MSFS." + lvardesignation + "_LOC_NAME", "");

                            TriggerCalcCode("0 (>L:SimConnect_" + lvardesignation + "LOCfreq)");
                            TriggerCalcCode("0 (>L:SimConnect_" + lvardesignation + "LOCheading)");

                            VoiceAttackPlugin.SetText("MSFS.FacilityReqVORICAO", "");
                            VoiceAttackPlugin.SetText("MSFS.FacilityReqVORregion", "");

                            Utils.scLocFreq = 0;
                            Utils.scLocHeading = 0;
                            Utils.scLocName = "";

                            switch (Utils.scPrimDesign1)
                            {
                                case 0:
                                    break;
                                case 1:
                                    pdesignator = "L";
                                    break;
                                case 2:
                                    pdesignator = "R";
                                    break;
                                case 3:
                                    pdesignator = "C";
                                    break;
                                case 4:
                                    pdesignator = "WATER";
                                    break;
                                case 5:
                                    pdesignator = "A";
                                    break;
                                case 6:
                                    pdesignator = "B";
                                    break;
                                case 7:
                                    pdesignator = "LAST";
                                    break;

                            }
                            switch (Utils.scSecDesign1)
                            {
                                case 0:
                                    break;
                                case 1:
                                    sdesignator = "L";
                                    break;
                                case 2:
                                    sdesignator = "R";
                                    break;
                                case 3:
                                    sdesignator = "C";
                                    break;
                                case 4:
                                    sdesignator = "WATER";
                                    break;
                                case 5:
                                    sdesignator = "A";
                                    break;
                                case 6:
                                    sdesignator = "B";
                                    break;
                                case 7:
                                    sdesignator = "LAST";
                                    break;

                            }



                            if (pdesignator != "NULL")
                            {
                                adesignation = Utils.scPrimNumb1.ToString() + pdesignator;
                            }
                            else
                            {
                                adesignation = Utils.scPrimNumb1.ToString();
                            }

                            if (sdesignator != "NULL")
                            {
                                bdesignation = Utils.scSecNumb1.ToString() + sdesignator;
                            }
                            else
                            {
                                bdesignation = Utils.scSecNumb1.ToString();
                            }


                            if (Utils.reqRunway == adesignation)
                            {
                                match = true;

                                VoiceAttackPlugin.LogOutput("Match: " + adesignation, "grey");

                                TriggerCalcCode(Utils.scLength1 + " (>L:SimConnect_" + lvardesignation + "Length)");
                                VoiceAttackPlugin.LogOutput("Runway " + adesignation + " length: " + Utils.scLength1, "blue");

                                TriggerCalcCode(Utils.scHeading1 + " (>L:SimConnect_" + lvardesignation + "Heading)");
                                VoiceAttackPlugin.LogOutput("Runway " + adesignation + " heading: " + Utils.scHeading1, "blue");

                                VoiceAttackPlugin.SetText("MSFS.FacilityReqVORICAO", Utils.scPrimVORICAO1);
                                VoiceAttackPlugin.SetText("MSFS.FacilityReqVORregion", Utils.scPrimVORregion1);
                                VoiceAttackPlugin.LogOutput("Saved VA variable VOR ICAO: " + Utils.scPrimVORICAO1 + " region: " + Utils.scPrimVORregion1, "grey");
                            }



                            else if (Utils.reqRunway == bdesignation)
                            {
                                match = true;

                                VoiceAttackPlugin.LogOutput("Match: " + bdesignation, "grey");

                                TriggerCalcCode(Utils.scLength1 + " (>L:SimConnect_" + lvardesignation + "Length)");
                                VoiceAttackPlugin.LogOutput("Runway " + bdesignation + " length: " + Utils.scLength1, "blue");

                                if (Utils.scHeading1 <= 180)
                                {
                                    util = Utils.scHeading1 + 180;
                                }
                                else
                                {
                                    util = Utils.scHeading1 - 180;
                                }

                                util = (float)Math.Round(util, 3);

                                TriggerCalcCode(util + " (>L:SimConnect_" + lvardesignation + "Heading)");
                                VoiceAttackPlugin.LogOutput("Runway " + bdesignation + " heading: " + util, "blue");

                                VoiceAttackPlugin.SetText("MSFS.FacilityReqVORICAO", Utils.scSecVORICAO1);
                                VoiceAttackPlugin.SetText("MSFS.FacilityReqVORregion", Utils.scSecVORregion1);
                                VoiceAttackPlugin.LogOutput("Saved VA variable VOR ICAO: " + Utils.scSecVORICAO1 + " region: " + Utils.scSecVORregion1, "grey");
                            }

                            switch (Utils.scPrimDesign2)
                            {
                                case 0:
                                    break;
                                case 1:
                                    pdesignator = "L";
                                    break;
                                case 2:
                                    pdesignator = "R";
                                    break;
                                case 3:
                                    pdesignator = "C";
                                    break;
                                case 4:
                                    pdesignator = "WATER";
                                    break;
                                case 5:
                                    pdesignator = "A";
                                    break;
                                case 6:
                                    pdesignator = "B";
                                    break;
                                case 7:
                                    pdesignator = "LAST";
                                    break;

                            }
                            switch (Utils.scSecDesign2)
                            {
                                case 0:
                                    break;
                                case 1:
                                    sdesignator = "L";
                                    break;
                                case 2:
                                    sdesignator = "R";
                                    break;
                                case 3:
                                    sdesignator = "C";
                                    break;
                                case 4:
                                    sdesignator = "WATER";
                                    break;
                                case 5:
                                    sdesignator = "A";
                                    break;
                                case 6:
                                    sdesignator = "B";
                                    break;
                                case 7:
                                    sdesignator = "LAST";
                                    break;
                            }


                            if (pdesignator != "NULL")
                            {
                                adesignation = Utils.scPrimNumb2.ToString() + pdesignator;
                            }
                            else
                            {
                                adesignation = Utils.scPrimNumb2.ToString();
                            }

                            if (sdesignator != "NULL")
                            {
                                bdesignation = Utils.scSecNumb2.ToString() + sdesignator;
                            }
                            else
                            {
                                bdesignation = Utils.scSecNumb2.ToString();
                            }


                            if (Utils.reqRunway == adesignation)
                            {
                                match = true;

                                VoiceAttackPlugin.LogOutput("Match: " + adesignation, "grey");

                                TriggerCalcCode(Utils.scLength2 + " (>L:SimConnect_" + lvardesignation + "Length)");
                                VoiceAttackPlugin.LogOutput("Runway " + adesignation + " length: " + Utils.scLength2, "blue");

                                TriggerCalcCode(Utils.scHeading2 + " (>L:SimConnect_" + lvardesignation + "Heading)");
                                VoiceAttackPlugin.LogOutput("Runway " + adesignation + " heading: " + Utils.scHeading2, "blue");

                                VoiceAttackPlugin.SetText("MSFS.FacilityReqVORICAO", Utils.scPrimVORICAO2);
                                VoiceAttackPlugin.SetText("MSFS.FacilityReqVORregion", Utils.scPrimVORregion2);
                                VoiceAttackPlugin.LogOutput("Saved VA variable VOR ICAO: " + Utils.scPrimVORICAO2 + " region: " + Utils.scPrimVORregion2, "grey");
                            }

                            else if (Utils.reqRunway == bdesignation)
                            {
                                match = true;

                                VoiceAttackPlugin.LogOutput("Match: " + bdesignation, "grey");

                                TriggerCalcCode(Utils.scLength2 + " (>L:SimConnect_" + lvardesignation + "Length)");
                                VoiceAttackPlugin.LogOutput("Runway " + bdesignation + " length: " + Utils.scLength2, "blue");

                                if (Utils.scHeading2 <= 180)
                                {
                                    util = Utils.scHeading2 + 180;
                                }
                                else
                                {
                                    util = Utils.scHeading2 - 180;
                                }

                                util = (float)Math.Round(util, 3);

                                TriggerCalcCode(util + " (>L:SimConnect_" + lvardesignation + "Heading)");
                                VoiceAttackPlugin.LogOutput("Runway " + bdesignation + " heading: " + util, "blue");

                                VoiceAttackPlugin.SetText("MSFS.FacilityReqVORICAO", Utils.scSecVORICAO2);
                                VoiceAttackPlugin.SetText("MSFS.FacilityReqVORregion", Utils.scSecVORregion2);
                                VoiceAttackPlugin.LogOutput("Saved VA variable VOR ICAO: " + Utils.scSecVORICAO2 + " region: " + Utils.scSecVORregion2, "grey");
                            }

                            switch (Utils.scPrimDesign3)
                            {
                                case 0:
                                    break;
                                case 1:
                                    pdesignator = "L";
                                    break;
                                case 2:
                                    pdesignator = "R";
                                    break;
                                case 3:
                                    pdesignator = "C";
                                    break;
                                case 4:
                                    pdesignator = "WATER";
                                    break;
                                case 5:
                                    pdesignator = "A";
                                    break;
                                case 6:
                                    pdesignator = "B";
                                    break;
                                case 7:
                                    pdesignator = "LAST";
                                    break;

                            }
                            switch (Utils.scSecDesign3)
                            {
                                case 0:
                                    break;
                                case 1:
                                    sdesignator = "L";
                                    break;
                                case 2:
                                    sdesignator = "R";
                                    break;
                                case 3:
                                    sdesignator = "C";
                                    break;
                                case 4:
                                    sdesignator = "WATER";
                                    break;
                                case 5:
                                    sdesignator = "A";
                                    break;
                                case 6:
                                    sdesignator = "B";
                                    break;
                                case 7:
                                    sdesignator = "LAST";
                                    break;

                            }


                            if (pdesignator != "NULL")
                            {
                                adesignation = Utils.scPrimNumb3.ToString() + pdesignator;
                            }
                            else
                            {
                                adesignation = Utils.scPrimNumb3.ToString();
                            }

                            if (sdesignator != "NULL")
                            {
                                bdesignation = Utils.scSecNumb3.ToString() + sdesignator;
                            }
                            else
                            {
                                bdesignation = Utils.scSecNumb3.ToString();
                            }


                            if (Utils.reqRunway == adesignation)
                            {
                                match = true;

                                VoiceAttackPlugin.LogOutput("Match: " + adesignation, "grey");

                                TriggerCalcCode(Utils.scLength3 + " (>L:SimConnect_" + lvardesignation + "Length)");
                                VoiceAttackPlugin.LogOutput("Runway " + adesignation + " length: " + Utils.scLength3, "blue");

                                TriggerCalcCode(Utils.scHeading3 + " (>L:SimConnect_" + lvardesignation + "Heading)");
                                VoiceAttackPlugin.LogOutput("Runway " + adesignation + " heading: " + Utils.scHeading3, "blue");

                                VoiceAttackPlugin.SetText("MSFS.FacilityReqVORICAO", Utils.scPrimVORICAO3);
                                VoiceAttackPlugin.SetText("MSFS.FacilityReqVORregion", Utils.scPrimVORregion3);
                                VoiceAttackPlugin.LogOutput("Saved VA variable VOR ICAO: " + Utils.scPrimVORICAO3 + " region: " + Utils.scPrimVORregion3, "grey");
                            }

                            else if (Utils.reqRunway == bdesignation)
                            {
                                match = true;

                                VoiceAttackPlugin.LogOutput("Match: " + bdesignation, "grey");

                                TriggerCalcCode(Utils.scLength3 + " (>L:SimConnect_" + lvardesignation + "Length)");
                                VoiceAttackPlugin.LogOutput("Runway " + bdesignation + " length: " + Utils.scLength3, "blue");

                                if (Utils.scHeading3 <= 180)
                                {
                                    util = Utils.scHeading3 + 180;
                                }
                                else
                                {
                                    util = Utils.scHeading3 - 180;
                                }

                                util = (float)Math.Round(util, 3);

                                TriggerCalcCode(util + " (>L:SimConnect_" + lvardesignation + "Heading)");
                                VoiceAttackPlugin.LogOutput("Runway " + bdesignation + " heading: " + util, "blue");

                                VoiceAttackPlugin.SetText("MSFS.FacilityReqVORICAO", Utils.scSecVORICAO3);
                                VoiceAttackPlugin.SetText("MSFS.FacilityReqVORregion", Utils.scSecVORregion3);
                                VoiceAttackPlugin.LogOutput("Saved VA variable VOR ICAO: " + Utils.scSecVORICAO3 + " region: " + Utils.scSecVORregion3, "grey");
                            }

                            switch (Utils.scPrimDesign4)
                            {
                                case 0:
                                    break;
                                case 1:
                                    pdesignator = "L";
                                    break;
                                case 2:
                                    pdesignator = "R";
                                    break;
                                case 3:
                                    pdesignator = "C";
                                    break;
                                case 4:
                                    pdesignator = "WATER";
                                    break;
                                case 5:
                                    pdesignator = "A";
                                    break;
                                case 6:
                                    pdesignator = "B";
                                    break;
                                case 7:
                                    pdesignator = "LAST";
                                    break;

                            }
                            switch (Utils.scSecDesign4)
                            {
                                case 0:
                                    break;
                                case 1:
                                    sdesignator = "L";
                                    break;
                                case 2:
                                    sdesignator = "R";
                                    break;
                                case 3:
                                    sdesignator = "C";
                                    break;
                                case 4:
                                    sdesignator = "WATER";
                                    break;
                                case 5:
                                    sdesignator = "A";
                                    break;
                                case 6:
                                    sdesignator = "B";
                                    break;
                                case 7:
                                    sdesignator = "LAST";
                                    break;

                            }

                            if (pdesignator != "NULL")
                            {
                                adesignation = Utils.scPrimNumb4.ToString() + pdesignator;
                            }
                            else
                            {
                                adesignation = Utils.scPrimNumb4.ToString();
                            }

                            if (sdesignator != "NULL")
                            {
                                bdesignation = Utils.scSecNumb4.ToString() + sdesignator;
                            }
                            else
                            {
                                bdesignation = Utils.scSecNumb4.ToString();
                            }


                            if (Utils.reqRunway == adesignation)
                            {
                                match = true;

                                VoiceAttackPlugin.LogOutput("Match: " + adesignation, "grey");

                                TriggerCalcCode(Utils.scLength4 + " (>L:SimConnect_" + lvardesignation + "Length)");
                                VoiceAttackPlugin.LogOutput("Runway " + adesignation + " length: " + Utils.scLength4, "blue");

                                TriggerCalcCode(Utils.scHeading4 + " (>L:SimConnect_" + lvardesignation + "Heading)");
                                VoiceAttackPlugin.LogOutput("Runway " + adesignation + " heading: " + Utils.scHeading4, "blue");

                                VoiceAttackPlugin.SetText("MSFS.FacilityReqVORICAO", Utils.scPrimVORICAO4);
                                VoiceAttackPlugin.SetText("MSFS.FacilityReqVORregion", Utils.scPrimVORregion4);
                                VoiceAttackPlugin.LogOutput("Saved VA variable VOR ICAO: " + Utils.scPrimVORICAO4 + " region: " + Utils.scPrimVORregion4, "grey");

                            }

                            else if (Utils.reqRunway == bdesignation)
                            {
                                match = true;

                                VoiceAttackPlugin.LogOutput("Match: " + bdesignation, "grey");

                                TriggerCalcCode(Utils.scLength4 + " (>L:SimConnect_" + lvardesignation + "Length)");
                                VoiceAttackPlugin.LogOutput("Runway " + bdesignation + " length: " + Utils.scLength4, "blue");

                                if (Utils.scHeading4 <= 180)
                                {
                                    util = Utils.scHeading4 + 180;
                                }
                                else
                                {
                                    util = Utils.scHeading4 - 180;
                                }

                                util = (float)Math.Round(util, 3);

                                TriggerCalcCode(util + " (>L:SimConnect_" + lvardesignation + "Heading)");
                                VoiceAttackPlugin.LogOutput("Runway " + bdesignation + " heading: " + util, "blue");

                                VoiceAttackPlugin.SetText("MSFS.FacilityReqVORICAO", Utils.scSecVORICAO4);
                                VoiceAttackPlugin.SetText("MSFS.FacilityReqVORregion", Utils.scSecVORregion4);
                                VoiceAttackPlugin.LogOutput("Saved VA variable VOR ICAO: " + Utils.scSecVORICAO4 + " region: " + Utils.scSecVORregion4, "grey");
                            }

                            switch (Utils.scPrimDesign5)
                            {
                                case 0:
                                    break;
                                case 1:
                                    pdesignator = "L";
                                    break;
                                case 2:
                                    pdesignator = "R";
                                    break;
                                case 3:
                                    pdesignator = "C";
                                    break;
                                case 4:
                                    pdesignator = "WATER";
                                    break;
                                case 5:
                                    pdesignator = "A";
                                    break;
                                case 6:
                                    pdesignator = "B";
                                    break;
                                case 7:
                                    pdesignator = "LAST";
                                    break;

                            }
                            switch (Utils.scSecDesign5)
                            {
                                case 0:
                                    break;
                                case 1:
                                    sdesignator = "L";
                                    break;
                                case 2:
                                    sdesignator = "R";
                                    break;
                                case 3:
                                    sdesignator = "C";
                                    break;
                                case 4:
                                    sdesignator = "WATER";
                                    break;
                                case 5:
                                    sdesignator = "A";
                                    break;
                                case 6:
                                    sdesignator = "B";
                                    break;
                                case 7:
                                    sdesignator = "LAST";
                                    break;

                            }

                            if (pdesignator != "NULL")
                            {
                                adesignation = Utils.scPrimNumb5.ToString() + pdesignator;
                            }
                            else
                            {
                                adesignation = Utils.scPrimNumb5.ToString();
                            }

                            if (sdesignator != "NULL")
                            {
                                bdesignation = Utils.scSecNumb5.ToString() + sdesignator;
                            }
                            else
                            {
                                bdesignation = Utils.scSecNumb5.ToString();
                            }


                            if (Utils.reqRunway == adesignation)
                            {
                                match = true;

                                VoiceAttackPlugin.LogOutput("Match: " + adesignation, "grey");

                                TriggerCalcCode(Utils.scLength5 + " (>L:SimConnect_" + lvardesignation + "Length)");
                                VoiceAttackPlugin.LogOutput("Runway " + adesignation + " length: " + Utils.scLength5, "blue");

                                TriggerCalcCode(Utils.scHeading5 + " (>L:SimConnect_" + lvardesignation + "Heading)");
                                VoiceAttackPlugin.LogOutput("Runway " + adesignation + " heading: " + Utils.scHeading5, "blue");

                                VoiceAttackPlugin.SetText("MSFS.FacilityReqVORICAO", Utils.scPrimVORICAO5);
                                VoiceAttackPlugin.SetText("MSFS.FacilityReqVORregion", Utils.scPrimVORregion5);
                                VoiceAttackPlugin.LogOutput("Saved VA variable VOR ICAO: " + Utils.scPrimVORICAO5 + " region: " + Utils.scPrimVORregion5, "grey");

                            }

                            else if (Utils.reqRunway == bdesignation)
                            {
                                match = true;

                                VoiceAttackPlugin.LogOutput("Match: " + bdesignation, "grey");

                                TriggerCalcCode(Utils.scLength5 + " (>L:SimConnect_" + lvardesignation + "Length)");
                                VoiceAttackPlugin.LogOutput("Runway " + bdesignation + " length: " + Utils.scLength5, "blue");

                                if (Utils.scHeading5 <= 180)
                                {
                                    util = Utils.scHeading5 + 180;
                                }
                                else
                                {
                                    util = Utils.scHeading5 - 180;
                                }

                                util = (float)Math.Round(util, 3);

                                TriggerCalcCode(util + " (>L:SimConnect_" + lvardesignation + "Heading)");
                                VoiceAttackPlugin.LogOutput("Runway " + bdesignation + " heading: " + util, "blue");

                                VoiceAttackPlugin.SetText("MSFS.FacilityReqVORICAO", Utils.scSecVORICAO5);
                                VoiceAttackPlugin.SetText("MSFS.FacilityReqVORregion", Utils.scSecVORregion5);
                                VoiceAttackPlugin.LogOutput("Saved VA variable VOR ICAO: " + Utils.scSecVORICAO5 + " region: " + Utils.scSecVORregion5, "grey");
                            }

                            switch (Utils.scPrimDesign6)
                            {
                                case 0:
                                    break;
                                case 1:
                                    pdesignator = "L";
                                    break;
                                case 2:
                                    pdesignator = "R";
                                    break;
                                case 3:
                                    pdesignator = "C";
                                    break;
                                case 4:
                                    pdesignator = "WATER";
                                    break;
                                case 5:
                                    pdesignator = "A";
                                    break;
                                case 6:
                                    pdesignator = "B";
                                    break;
                                case 7:
                                    pdesignator = "LAST";
                                    break;

                            }
                            switch (Utils.scSecDesign6)
                            {
                                case 0:
                                    break;
                                case 1:
                                    sdesignator = "L";
                                    break;
                                case 2:
                                    sdesignator = "R";
                                    break;
                                case 3:
                                    sdesignator = "C";
                                    break;
                                case 4:
                                    sdesignator = "WATER";
                                    break;
                                case 5:
                                    sdesignator = "A";
                                    break;
                                case 6:
                                    sdesignator = "B";
                                    break;
                                case 7:
                                    sdesignator = "LAST";
                                    break;

                            }

                            if (pdesignator != "NULL")
                            {
                                adesignation = Utils.scPrimNumb6.ToString() + pdesignator;
                            }
                            else
                            {
                                adesignation = Utils.scPrimNumb6.ToString();
                            }

                            if (sdesignator != "NULL")
                            {
                                bdesignation = Utils.scSecNumb6.ToString() + sdesignator;
                            }
                            else
                            {
                                bdesignation = Utils.scSecNumb6.ToString();
                            }


                            if (Utils.reqRunway == adesignation)
                            {
                                match = true;

                                VoiceAttackPlugin.LogOutput("Match: " + adesignation, "grey");

                                TriggerCalcCode(Utils.scLength6 + " (>L:SimConnect_" + lvardesignation + "Length)");
                                VoiceAttackPlugin.LogOutput("Runway " + adesignation + " length: " + Utils.scLength6, "blue");

                                TriggerCalcCode(Utils.scHeading6 + " (>L:SimConnect_" + lvardesignation + "Heading)");
                                VoiceAttackPlugin.LogOutput("Runway " + adesignation + " heading: " + Utils.scHeading6, "blue");

                                VoiceAttackPlugin.SetText("MSFS.FacilityReqVORICAO", Utils.scPrimVORICAO6);
                                VoiceAttackPlugin.SetText("MSFS.FacilityReqVORregion", Utils.scPrimVORregion6);
                                VoiceAttackPlugin.LogOutput("Saved VA variable VOR ICAO: " + Utils.scPrimVORICAO6 + " region: " + Utils.scPrimVORregion6, "grey");

                            }

                            else if (Utils.reqRunway == bdesignation)
                            {
                                match = true;

                                VoiceAttackPlugin.LogOutput("Match: " + bdesignation, "grey");

                                TriggerCalcCode(Utils.scLength6 + " (>L:SimConnect_" + lvardesignation + "Length)");
                                VoiceAttackPlugin.LogOutput("Runway " + bdesignation + " length: " + Utils.scLength6, "blue");

                                if (Utils.scHeading6 <= 180)
                                {
                                    util = Utils.scHeading6 + 180;
                                }
                                else
                                {
                                    util = Utils.scHeading6 - 180;
                                }

                                util = (float)Math.Round(util, 3);

                                TriggerCalcCode(util + " (>L:SimConnect_" + lvardesignation + "Heading)");
                                VoiceAttackPlugin.LogOutput("Runway " + bdesignation + " heading: " + util, "blue");

                                VoiceAttackPlugin.SetText("MSFS.FacilityReqVORICAO", Utils.scSecVORICAO6);
                                VoiceAttackPlugin.SetText("MSFS.FacilityReqVORregion", Utils.scSecVORregion6);
                                VoiceAttackPlugin.LogOutput("Saved VA variable VOR ICAO: " + Utils.scSecVORICAO6 + " region: " + Utils.scSecVORregion6, "grey");
                            }

                            switch (Utils.scPrimDesign7)
                            {
                                case 0:
                                    break;
                                case 1:
                                    pdesignator = "L";
                                    break;
                                case 2:
                                    pdesignator = "R";
                                    break;
                                case 3:
                                    pdesignator = "C";
                                    break;
                                case 4:
                                    pdesignator = "WATER";
                                    break;
                                case 5:
                                    pdesignator = "A";
                                    break;
                                case 6:
                                    pdesignator = "B";
                                    break;
                                case 7:
                                    pdesignator = "LAST";
                                    break;

                            }
                            switch (Utils.scSecDesign7)
                            {
                                case 0:
                                    break;
                                case 1:
                                    sdesignator = "L";
                                    break;
                                case 2:
                                    sdesignator = "R";
                                    break;
                                case 3:
                                    sdesignator = "C";
                                    break;
                                case 4:
                                    sdesignator = "WATER";
                                    break;
                                case 5:
                                    sdesignator = "A";
                                    break;
                                case 6:
                                    sdesignator = "B";
                                    break;
                                case 7:
                                    sdesignator = "LAST";
                                    break;

                            }

                            if (pdesignator != "NULL")
                            {
                                adesignation = Utils.scPrimNumb7.ToString() + pdesignator;
                            }
                            else
                            {
                                adesignation = Utils.scPrimNumb7.ToString();
                            }

                            if (sdesignator != "NULL")
                            {
                                bdesignation = Utils.scSecNumb7.ToString() + sdesignator;
                            }
                            else
                            {
                                bdesignation = Utils.scSecNumb7.ToString();
                            }


                            if (Utils.reqRunway == adesignation)
                            {
                                match = true;

                                VoiceAttackPlugin.LogOutput("Match: " + adesignation, "grey");

                                TriggerCalcCode(Utils.scLength7 + " (>L:SimConnect_" + lvardesignation + "Length)");
                                VoiceAttackPlugin.LogOutput("Runway " + adesignation + " length: " + Utils.scLength7, "blue");

                                TriggerCalcCode(Utils.scHeading7 + " (>L:SimConnect_" + lvardesignation + "Heading)");
                                VoiceAttackPlugin.LogOutput("Runway " + adesignation + " heading: " + Utils.scHeading7, "blue");

                                VoiceAttackPlugin.SetText("MSFS.FacilityReqVORICAO", Utils.scPrimVORICAO7);
                                VoiceAttackPlugin.SetText("MSFS.FacilityReqVORregion", Utils.scPrimVORregion7);
                                VoiceAttackPlugin.LogOutput("Saved VA variable VOR ICAO: " + Utils.scPrimVORICAO7 + " region: " + Utils.scPrimVORregion7, "grey");

                            }

                            else if (Utils.reqRunway == bdesignation)
                            {
                                match = true;

                                VoiceAttackPlugin.LogOutput("Match: " + bdesignation, "grey");

                                TriggerCalcCode(Utils.scLength7 + " (>L:SimConnect_" + lvardesignation + "Length)");
                                VoiceAttackPlugin.LogOutput("Runway " + bdesignation + " length: " + Utils.scLength7, "blue");

                                if (Utils.scHeading7 <= 180)
                                {
                                    util = Utils.scHeading7 + 180;
                                }
                                else
                                {
                                    util = Utils.scHeading7 - 180;
                                }

                                util = (float)Math.Round(util, 3);

                                TriggerCalcCode(util + " (>L:SimConnect_" + lvardesignation + "Heading)");
                                VoiceAttackPlugin.LogOutput("Runway " + bdesignation + " heading: " + util, "blue");

                                VoiceAttackPlugin.SetText("MSFS.FacilityReqVORICAO", Utils.scSecVORICAO7);
                                VoiceAttackPlugin.SetText("MSFS.FacilityReqVORregion", Utils.scSecVORregion7);
                                VoiceAttackPlugin.LogOutput("Saved VA variable VOR ICAO: " + Utils.scSecVORICAO7 + " region: " + Utils.scSecVORregion7, "grey");
                            }

                            switch (Utils.scPrimDesign8)
                            {
                                case 0:
                                    break;
                                case 1:
                                    pdesignator = "L";
                                    break;
                                case 2:
                                    pdesignator = "R";
                                    break;
                                case 3:
                                    pdesignator = "C";
                                    break;
                                case 4:
                                    pdesignator = "WATER";
                                    break;
                                case 5:
                                    pdesignator = "A";
                                    break;
                                case 6:
                                    pdesignator = "B";
                                    break;
                                case 7:
                                    pdesignator = "LAST";
                                    break;

                            }
                            switch (Utils.scSecDesign8)
                            {
                                case 0:
                                    break;
                                case 1:
                                    sdesignator = "L";
                                    break;
                                case 2:
                                    sdesignator = "R";
                                    break;
                                case 3:
                                    sdesignator = "C";
                                    break;
                                case 4:
                                    sdesignator = "WATER";
                                    break;
                                case 5:
                                    sdesignator = "A";
                                    break;
                                case 6:
                                    sdesignator = "B";
                                    break;
                                case 7:
                                    sdesignator = "LAST";
                                    break;

                            }

                            if (pdesignator != "NULL")
                            {
                                adesignation = Utils.scPrimNumb8.ToString() + pdesignator;
                            }
                            else
                            {
                                adesignation = Utils.scPrimNumb8.ToString();
                            }

                            if (sdesignator != "NULL")
                            {
                                bdesignation = Utils.scSecNumb8.ToString() + sdesignator;
                            }
                            else
                            {
                                bdesignation = Utils.scSecNumb8.ToString();
                            }


                            if (Utils.reqRunway == adesignation)
                            {
                                match = true;

                                VoiceAttackPlugin.LogOutput("Match: " + adesignation, "grey");

                                TriggerCalcCode(Utils.scLength8 + " (>L:SimConnect_" + lvardesignation + "Length)");
                                VoiceAttackPlugin.LogOutput("Runway " + adesignation + " length: " + Utils.scLength8, "blue");

                                TriggerCalcCode(Utils.scHeading8 + " (>L:SimConnect_" + lvardesignation + "Heading)");
                                VoiceAttackPlugin.LogOutput("Runway " + adesignation + " heading: " + Utils.scHeading8, "blue");

                                VoiceAttackPlugin.SetText("MSFS.FacilityReqVORICAO", Utils.scPrimVORICAO8);
                                VoiceAttackPlugin.SetText("MSFS.FacilityReqVORregion", Utils.scPrimVORregion8);
                                VoiceAttackPlugin.LogOutput("Saved VA variable VOR ICAO: " + Utils.scPrimVORICAO8 + " region: " + Utils.scPrimVORregion8, "grey");

                            }

                            else if (Utils.reqRunway == bdesignation)
                            {
                                match = true;

                                VoiceAttackPlugin.LogOutput("Match: " + bdesignation, "grey");

                                TriggerCalcCode(Utils.scLength8 + " (>L:SimConnect_" + lvardesignation + "Length)");
                                VoiceAttackPlugin.LogOutput("Runway " + bdesignation + " length: " + Utils.scLength8, "blue");

                                if (Utils.scHeading8 <= 180)
                                {
                                    util = Utils.scHeading8 + 180;
                                }
                                else
                                {
                                    util = Utils.scHeading8 - 180;
                                }

                                util = (float)Math.Round(util, 3);

                                TriggerCalcCode(util + " (>L:SimConnect_" + lvardesignation + "Heading)");
                                VoiceAttackPlugin.LogOutput("Runway " + bdesignation + " heading: " + util, "blue");

                                VoiceAttackPlugin.SetText("MSFS.FacilityReqVORICAO", Utils.scSecVORICAO8);
                                VoiceAttackPlugin.SetText("MSFS.FacilityReqVORregion", Utils.scSecVORregion8);
                                VoiceAttackPlugin.LogOutput("Saved VA variable VOR ICAO: " + Utils.scSecVORICAO8 + " region: " + Utils.scSecVORregion8, "grey");
                            }

                        }




                        catch (Exception ex)
                        {
                            VoiceAttackPlugin.LogErrorOutput("Wrapping Up problem with RUNWAY type. " + ex, "red");
                        }
                        break;

                    case "VOR":

                        try
                        {
                            lvardesignation = VoiceAttackPlugin.GetText("MSFS.LvarVorDes");

                            TriggerCalcCode(Utils.scLocFreq + " (>L:SimConnect_" + lvardesignation + "LOCfreq)");
                            VoiceAttackPlugin.LogOutput("Localizer Frequency: " + Utils.scLocFreq + "MHz", "blue");
                            VoiceAttackPlugin.SetText("MSFS." + lvardesignation + "_LOC_FREQ", Utils.scLocFreq.ToString());

                            TriggerCalcCode(Utils.scLocHeading + " (>L:SimConnect_" + lvardesignation + "LOCheading)");
                            VoiceAttackPlugin.LogOutput("Localizer Heading: " + Utils.scLocHeading, "blue");
                            VoiceAttackPlugin.SetText("MSFS." + lvardesignation + "_LOC_HEAD", Utils.scLocHeading.ToString());

                            VoiceAttackPlugin.SetText("MSFS." + lvardesignation + "_LOC_NAME", Utils.scLocName);
                        }
                        catch (Exception ex)
                        {
                            VoiceAttackPlugin.LogErrorOutput("Wrapping Up problem with VOR type. " + ex, "red");
                        }


                        break;
                }
            }
            catch (Exception ex)
            {

                VoiceAttackPlugin.LogOutput("FacilityWrappingUp error:" + ex, "grey");
            }



        }
        public void TriggerKeySimconnect(FsControlList varName, string varData = "0")
        {

            VoiceAttackPlugin.LogOutput("Triggering Key Event method through SimConnect.", "grey");

            UInt32 kData;
            Decimal d;

            if (String.IsNullOrWhiteSpace(varData)) varData = "0";

            try
            {                                

                Byte[] Bytes = BitConverter.GetBytes(Convert.ToInt32(varData));
                kData = BitConverter.ToUInt32(Bytes, 0);
                _simConnection.TransmitClientEvent(SimConnect.SIMCONNECT_OBJECT_ID_USER, varName, kData, NOTIFICATION_GROUPS.DEFAULT, SIMCONNECT_EVENT_FLAG.GROUPID_IS_PRIORITY);
                VoiceAttackPlugin.LogOutput("Data sent.", "grey");


            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogOutput("Failed to send key event data through SimConnect. " + e.Message, "red");

            }

            


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

                        try
                        {
                            _waSimConnection.executeCalculatorCode(calcCode, CalcResultType.Double, out double fResult, out string sResult);
                            VoiceAttackPlugin.LogOutput("Executed calculator code with 3 index digits: " + calcCode, "grey");
                            if (varName == "TRANSPONDER CODE")
                            {

                                fResult = fResult * 1000000;

                            }

                            return fResult;

                        }
                        catch (Exception e)
                        {
                            VoiceAttackPlugin.LogErrorOutput("Failed to receive SimVar data. " + e.Message, "red");
                            return 0;
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

                        string calcCode = "(A:" + varName + ":" + blast + last + ", " + unit + ")";
                        try
                        {
                            _waSimConnection.executeCalculatorCode(calcCode, CalcResultType.Double, out double fResult, out string sResult);
                            VoiceAttackPlugin.LogOutput("Executed calculator code with 2 index digits: " + calcCode, "grey");
                            if (varName == "TRANSPONDER CODE")
                            {

                                fResult = fResult * 1000000;

                            }

                            return fResult;

                        }
                        catch (Exception e)
                        {
                            VoiceAttackPlugin.LogErrorOutput("Failed to receive SimVar data. " + e.Message, "red");
                            return 0;
                        }

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

                    try
                    {
                        _waSimConnection.executeCalculatorCode(calcCode, CalcResultType.Double, out double fResult, out string sResult);
                        VoiceAttackPlugin.LogOutput("Executed calculator code with 1 index digits: " + calcCode, "grey");
                        if (varName == "TRANSPONDER CODE")
                        {

                            fResult = fResult * 1000000;

                        }

                        return fResult;

                    }
                    catch (Exception e)
                    {
                        VoiceAttackPlugin.LogErrorOutput("Failed to receive SimVar data. " + e.Message, "red");
                        return 0;
                    }


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

                try
                {
                    _waSimConnection.executeCalculatorCode(calcCode, CalcResultType.Double, out double fResult, out string sResult);
                    VoiceAttackPlugin.LogOutput("Executed calculator: " + calcCode, "grey");
                    if (varName == "TRANSPONDER CODE")
                    {

                        fResult = fResult * 1000000;

                    }

                    return fResult;

                }
                catch (Exception e)
                {
                    VoiceAttackPlugin.LogErrorOutput("Failed to receive SimVar data. " + e.Message, "red");
                    return 0;
                }

            }

        }

        public void GetSimVarSimConnect_2()
        {                     

            {
                VoiceAttackPlugin.LogOutput("Connecting SimConnect...", "grey");

                try
                {
                    int count = 0;

                    _simConnection = new SimConnect("FVCplugin", IntPtr.Zero, WM_USER_SIMCONNECT, null, 0);

                    /// Listen to connect and quit msgs
                    _simConnection.OnRecvOpen += simconnect_OnRecvOpen;
                    _simConnection.OnRecvQuit += simconnect_OnRecvQuit;
                    //_simConnection.OnRecvException += simconnect_OnRecvException;

                    _simConnection.OnRecvSimobjectDataBytype += simconnect_OnRecvSimobjectDataBytype;

                    _simConnection.SetNotificationGroupPriority(NOTIFICATION_GROUPS.DEFAULT, SimConnect.SIMCONNECT_GROUP_PRIORITY_HIGHEST);

                    VoiceAttackPlugin.LogOutput("Connection established", "grey");

                    while (isLastMessageReceived == false && count <= 10)
                    {
                        count ++;

                        System.Threading.Thread.Sleep(100);
                        
                        _simConnection.ReceiveMessage();

                    }


                }
                catch (Exception ex)
                {

                    VoiceAttackPlugin.LogOutput("" + ex, "grey");
                }



            }

        }

        public void GetSimVarSimConnect_1(string varName)
        {

            requestType = "SIMVAR";

            VoiceAttackPlugin.LogOutput("Triggering SimVar SimConnect read method...", "grey");

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

                        VoiceAttackPlugin.LogOutput("SimConnect type: " + Utils.dataType, "grey");
                        VoiceAttackPlugin.LogOutput("Result data type: " + Utils.resultDataType, "grey");

                        Utils.simvar = varName + ":" + clast + blast + last;
                        Utils.simvarindex = "" + clast + blast + last;
                        Utils.simvarunit = unit;

                        VoiceAttackPlugin.LogOutput("3 index digit SimVar registered: " + Utils.simvar, "grey");

                        GetSimVarSimConnect_2();

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

                        VoiceAttackPlugin.LogOutput("SimConnect type: " + Utils.dataType, "grey");
                        VoiceAttackPlugin.LogOutput("Result data type: " + Utils.resultDataType, "grey");

                        Utils.simvar = varName + ":" + blast + last;
                        Utils.simvarindex = "" + blast + last;
                        Utils.simvarunit = unit;

                        VoiceAttackPlugin.LogOutput("2 index digit SimVar registered: " + Utils.simvar, "grey");

                        do
                        {
                            GetSimVarSimConnect_2();
                        }

                        while (isLastMessageReceived == false);
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

                    VoiceAttackPlugin.LogOutput("SimConnect type: " + Utils.dataType, "grey");
                    VoiceAttackPlugin.LogOutput("Result data type: " + Utils.resultDataType, "grey");

                    Utils.simvar = varName + ":" + last;
                    Utils.simvarindex = "" + last;
                    Utils.simvarunit = unit;

                    VoiceAttackPlugin.LogOutput("1 index digit SimVar registered: " + Utils.simvar, "grey");
                    
                    GetSimVarSimConnect_2();
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

                VoiceAttackPlugin.LogOutput("SimConnect type: " + Utils.dataType, "grey");
                VoiceAttackPlugin.LogOutput("Result data type: " + Utils.resultDataType, "grey");

                Utils.simvar = varName;
                Utils.simvarunit = unit;

                VoiceAttackPlugin.LogOutput("0 index digit SimVar registered: " + Utils.simvar, "grey");
                
                GetSimVarSimConnect_2();
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
            VoiceAttackPlugin.LogMonitorOutput("CostIndex: " + Utils.sbCostIndex, "blue");

            TriggerCalcCode(Utils.sbInitialAlt + " (>L:SimBrief_InitialAlt)");
            VoiceAttackPlugin.LogMonitorOutput("InitialAlt: " + Utils.sbInitialAlt, "blue");

            TriggerCalcCode(Utils.sbAvgWindDir + " (>L:SimBrief_AvgWindDir)");
            VoiceAttackPlugin.LogMonitorOutput("AvgWindDir: " + Utils.sbAvgWindDir, "blue");

            TriggerCalcCode(Utils.sbAvgWindSpd + " (>L:SimBrief_AvgWindSpd)");
            VoiceAttackPlugin.LogMonitorOutput("AvgWindSpd: " + Utils.sbAvgWindSpd, "blue");

            TriggerCalcCode(Utils.sbTopClimbOAT + " (>L:SimBrief_TopClimbOAT)");
            VoiceAttackPlugin.LogMonitorOutput("TopClimbOAT: " + Utils.sbTopClimbOAT, "blue");

            TriggerCalcCode(Utils.sbOriginElevation + " (>L:SimBrief_OriginElevation)");
            VoiceAttackPlugin.LogMonitorOutput("OriginElevation: " + Utils.sbOriginElevation, "blue");

            TriggerCalcCode(Utils.sbOriginTransAlt + " (>L:SimBrief_OriginTransAlt)");
            VoiceAttackPlugin.LogMonitorOutput("OriginTransAlt: " + Utils.sbOriginTransAlt, "blue");

            TriggerCalcCode(Utils.sbOriginTransLevel + " (>L:SimBrief_OriginTransLevel)");
            VoiceAttackPlugin.LogMonitorOutput("OriginTransLevel: " + Utils.sbOriginTransLevel, "blue");

            TriggerCalcCode(Utils.sbAltnElevation + " (>L:SimBrief_AltnElevation)");
            VoiceAttackPlugin.LogMonitorOutput("AltnElevation: " + Utils.sbAltnElevation, "blue");

            TriggerCalcCode(Utils.sbAltnTransAlt + " (>L:SimBrief_AltnTransAlt)");
            VoiceAttackPlugin.LogMonitorOutput("AltnTransAlt: " + Utils.sbAltnTransAlt, "blue");

            TriggerCalcCode(Utils.sbAltnTransLevel + " (>L:SimBrief_AltnTransLevel)");
            VoiceAttackPlugin.LogMonitorOutput("AltnTransLevel: " + Utils.sbAltnTransLevel, "blue");

            TriggerCalcCode(Utils.sbDestElevation + " (>L:SimBrief_DestElevation)");
            VoiceAttackPlugin.LogMonitorOutput("DestElevation: " + Utils.sbDestElevation, "blue");

            TriggerCalcCode(Utils.sbDestTransAlt + " (>L:SimBrief_DestTransAlt)");
            VoiceAttackPlugin.LogMonitorOutput("DestTransAlt: " + Utils.sbDestTransAlt, "blue");

            TriggerCalcCode(Utils.sbDestTransLevel + " (>L:SimBrief_DestTransLevel)");
            VoiceAttackPlugin.LogMonitorOutput("DestTransLevel: " + Utils.sbDestTransLevel, "blue");

            TriggerCalcCode(Utils.sbFinRes + " (>L:SimBrief_FinRes)");
            VoiceAttackPlugin.LogMonitorOutput("FinRes: " + Utils.sbFinRes, "blue");

            TriggerCalcCode(Utils.sbAltnFuel + " (>L:SimBrief_AltnFuel)");
            VoiceAttackPlugin.LogMonitorOutput("AltnFuel: " + Utils.sbAltnFuel, "blue");

            TriggerCalcCode(Utils.sbFinresPAltn + " (>L:SimBrief_FinresPAltn)");
            VoiceAttackPlugin.LogMonitorOutput("FinresPAltn: " + Utils.sbFinresPAltn, "blue");

            TriggerCalcCode(Utils.sbFuel + " (>L:SimBrief_Fuel)");
            VoiceAttackPlugin.LogMonitorOutput("Fuel: " + Utils.sbFuel, "blue");

            TriggerCalcCode(Utils.sbPassenger + " (>L:SimBrief_Passenger)");
            VoiceAttackPlugin.LogMonitorOutput("Passenger: " + Utils.sbPassenger, "blue");

            TriggerCalcCode(Utils.sbBags + " (>L:SimBrief_Bags)");
            VoiceAttackPlugin.LogMonitorOutput("Bags: " + Utils.sbBags, "blue");

            TriggerCalcCode(Utils.sbWeightPax + " (>L:SimBrief_WeightPax)");
            VoiceAttackPlugin.LogMonitorOutput("WeightPax: " + Utils.sbWeightPax, "blue");

            TriggerCalcCode(Utils.sbWeightCargo + " (>L:SimBrief_WeightCargo)");
            VoiceAttackPlugin.LogMonitorOutput("WeightCargo: " + Utils.sbWeightCargo, "blue");

            TriggerCalcCode(Utils.sbPayload + " (>L:SimBrief_Payload)");
            VoiceAttackPlugin.LogMonitorOutput("Payload: " + Utils.sbPayload, "blue");

            TriggerCalcCode(Utils.sbZFW + " (>L:SimBrief_ZFW)");
            VoiceAttackPlugin.LogMonitorOutput("ZFW: " + Utils.sbZFW, "blue");

            TriggerCalcCode(Utils.sbTOW + " (>L:SimBrief_TOW)");
            VoiceAttackPlugin.LogMonitorOutput("TOW: " + Utils.sbTOW, "blue");

            TriggerCalcCode(Utils.sbDestAvailWeather + " (>L:SimBrief_DestWeatherSource)");
            VoiceAttackPlugin.LogMonitorOutput("DestWeatherSource: " + Utils.sbDestAvailWeather, "blue");

            TriggerCalcCode(Utils.sbOriginAvailWeather + " (>L:SimBrief_OriginWeatherSource)");
            VoiceAttackPlugin.LogMonitorOutput("OriginWeatherSource: " + Utils.sbOriginAvailWeather, "blue");

            TriggerCalcCode(Utils.sbAltnAvailWeather + " (>L:SimBrief_AltnWeatherSource)");
            VoiceAttackPlugin.LogMonitorOutput("AltnWeatherSource: " + Utils.sbAltnAvailWeather, "blue");

            VoiceAttackPlugin.SetText("sbFlight", Utils.sbFlight);
            VoiceAttackPlugin.LogMonitorOutput("sbFlight: " + Utils.sbFlight, "blue");

            VoiceAttackPlugin.SetText("sbAirlineICAO", Utils.sbAirlineICAO);
            VoiceAttackPlugin.LogMonitorOutput("sbAirlineICAO: " + Utils.sbAirlineICAO, "blue");

            Utils.SetCallsign();

            VoiceAttackPlugin.SetText("sbCallsign", Utils.sbCallsign);
            VoiceAttackPlugin.LogMonitorOutput("sbCallsign: " + Utils.sbCallsign, "blue");

            VoiceAttackPlugin.SetText("sbCruiseProf", Utils.sbCruiseProf);
            VoiceAttackPlugin.LogMonitorOutput("sbCruiseProf: " + Utils.sbCruiseProf, "blue");
            
            VoiceAttackPlugin.SetText("sbClimbProf", Utils.sbClimbProf);
            VoiceAttackPlugin.LogMonitorOutput("sbClimbProf: " + Utils.sbClimbProf, "blue");

            VoiceAttackPlugin.SetText("sbDescentProf", Utils.sbDescentProf);
            VoiceAttackPlugin.LogMonitorOutput("sbDescentProf: " + Utils.sbDescentProf, "blue");

            VoiceAttackPlugin.SetText("sbRoute", Utils.sbRoute);
            VoiceAttackPlugin.LogMonitorOutput("sbRoute: " + Utils.sbRoute, "blue");

            VoiceAttackPlugin.SetText("sbOrigin", Utils.sbOrigin);
            VoiceAttackPlugin.LogMonitorOutput("sbOrigin: " + Utils.sbOrigin, "blue");

            VoiceAttackPlugin.SetText("sbOriginRwy", Utils.sbOriginRwy);
            VoiceAttackPlugin.LogMonitorOutput("sbOriginRwy: " + Utils.sbOriginRwy, "blue");

            VoiceAttackPlugin.SetText("sbOriginMetar", Utils.sbOriginMetar);
            VoiceAttackPlugin.LogMonitorOutput("sbOriginMetar: " + Utils.sbOriginMetar, "blue");

            VoiceAttackPlugin.SetText("sbOriginTAF", Utils.sbOriginTAF);
            VoiceAttackPlugin.LogMonitorOutput("sbOriginTAF: " + Utils.sbOriginTAF, "blue");

            VoiceAttackPlugin.SetText("sbOriginWind", Utils.sbOriginWind);
            VoiceAttackPlugin.LogMonitorOutput("sbOriginWind: " + Utils.sbOriginWind, "blue");

            VoiceAttackPlugin.SetText("sbOriginPressure", Utils.sbOriginPressure);
            VoiceAttackPlugin.LogMonitorOutput("sbOriginPressure: " + Utils.sbOriginPressure, "blue");

            VoiceAttackPlugin.SetText("sbOriginTemp", Utils.sbOriginTemp);
            VoiceAttackPlugin.LogMonitorOutput("sbOriginTemp: " + Utils.sbOriginTemp, "blue");

            VoiceAttackPlugin.SetText("sbAltn", Utils.sbAltn);
            VoiceAttackPlugin.LogMonitorOutput("sbAltn: " + Utils.sbAltn, "blue");

            VoiceAttackPlugin.SetText("sbAltnRwy", Utils.sbAltnRwy);
            VoiceAttackPlugin.LogMonitorOutput("sbAltnRwy: " + Utils.sbAltnRwy, "blue");

            VoiceAttackPlugin.SetText("sbAltnMetar", Utils.sbAltnMetar);
            VoiceAttackPlugin.LogMonitorOutput("sbAltnMetar: " + Utils.sbAltnMetar, "blue");

            VoiceAttackPlugin.SetText("sbAltnTAF", Utils.sbAltnTAF);
            VoiceAttackPlugin.LogMonitorOutput("sbAltnTAF: " + Utils.sbAltnTAF, "blue");

            VoiceAttackPlugin.SetText("sbAltnWind", Utils.sbAltnWind);
            VoiceAttackPlugin.LogMonitorOutput("sbAltnWind: " + Utils.sbAltnWind, "blue");

            VoiceAttackPlugin.SetText("sbAltnPressure", Utils.sbAltnPressure);
            VoiceAttackPlugin.LogMonitorOutput("sbAltnPressure: " + Utils.sbAltnPressure, "blue");

            VoiceAttackPlugin.SetText("sbAltnTemp", Utils.sbAltnTemp);
            VoiceAttackPlugin.LogMonitorOutput("sbAltnTemp: " + Utils.sbAltnTemp, "blue");

            VoiceAttackPlugin.SetText("sbDestination", Utils.sbDestination);
            VoiceAttackPlugin.LogMonitorOutput("sbDestination: " + Utils.sbDestination, "blue");

            VoiceAttackPlugin.SetText("sbDestRwy", Utils.sbDestRwy);
            VoiceAttackPlugin.LogMonitorOutput("sbDestRwy: " + Utils.sbDestRwy, "blue");

            VoiceAttackPlugin.SetText("sbDestMetar", Utils.sbDestMetar);
            VoiceAttackPlugin.LogMonitorOutput("sbDestMetar: " + Utils.sbDestMetar, "blue");

            VoiceAttackPlugin.SetText("sbDestTAF", Utils.sbDestTAF);
            VoiceAttackPlugin.LogMonitorOutput("sbDestTAF: " + Utils.sbDestTAF, "blue");

            VoiceAttackPlugin.SetText("sbDestWind", Utils.sbDestWind);
            VoiceAttackPlugin.LogMonitorOutput("sbDestWind: " + Utils.sbDestWind, "blue");

            VoiceAttackPlugin.SetText("sbDestPressure", Utils.sbDestPressure);
            VoiceAttackPlugin.LogMonitorOutput("sbDestPressure: " + Utils.sbDestPressure, "blue");

            VoiceAttackPlugin.SetText("sbDestTemp", Utils.sbDestTemp);
            VoiceAttackPlugin.LogMonitorOutput("sbDestTemp: " + Utils.sbDestTemp, "blue");

            VoiceAttackPlugin.SetText("sbUnits", Utils.sbUnits);
            VoiceAttackPlugin.LogMonitorOutput("sbUnits: " + Utils.sbUnits, "blue");




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
            try
            {
                _waSimConnection.OnClientEvent += ClientStatusHandler;
                _waSimConnection.OnLogRecordReceived += LogHandler;
                _waSimConnection.OnDataReceived += DataSubscriptionHandler;
                _waSimConnection.setLogLevel(LogLevel.Trace, LogFacility.File, LogSource.Client);
                _waSimConnection.setLogLevel(LogLevel.Trace, LogFacility.Remote, LogSource.Client);
                _waSimConnection.setLogLevel(LogLevel.Trace, LogFacility.File, LogSource.Server);
                _waSimConnection.setLogLevel(LogLevel.Trace, LogFacility.Remote, LogSource.Server);

            }
            catch (Exception e)
            {
                VoiceAttackPlugin.LogErrorOutput("Failed to initialize WASM handlers. " + e.Message, "red");
            }


        }

        /// <summary>
        /// Catches any exceptions that are encountered by SimConnect
        /// </summary>
        private void simconnect_OnRecvException(SimConnect sender, SIMCONNECT_RECV_EXCEPTION data)
        {
            SIMCONNECT_EXCEPTION e = (SIMCONNECT_EXCEPTION)data.dwException;
            VoiceAttackPlugin.LogErrorOutput("SimConnect_OnRecvException: " + e.ToString(), "red");

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

            VoiceAttackPlugin.LogOutput(System.Reflection.MethodBase.GetCurrentMethod()?.Name, "grey");

            ++currentRequest;
            REQUEST_DEFINITON rd = (REQUEST_DEFINITON)currentRequest;

            ++currentDefinition;
            SIMVAR_DEFINITION sd = (SIMVAR_DEFINITION)currentDefinition;

            switch (requestType)
            {
                case "AIRPORT":

                    VoiceAttackPlugin.LogOutput("OnRecvOpen for AIRPORT request", "grey");

                    _simConnection.AddToFacilityDefinition(sd, "OPEN AIRPORT");

                    _simConnection.AddToFacilityDefinition(sd, "LATITUDE");
                    _simConnection.AddToFacilityDefinition(sd, "N_ARRIVALS");
                    _simConnection.AddToFacilityDefinition(sd, "N_RUNWAYS");


                    _simConnection.AddToFacilityDefinition(sd, "CLOSE AIRPORT");

                    _simConnection.RegisterFacilityDataDefineStruct<airport>(SIMCONNECT_FACILITY_DATA_TYPE.AIRPORT);
                    _simConnection.RequestFacilityData(sd, rd, "SCEL", "");

                    break;

                case "RUNWAY":

                    VoiceAttackPlugin.LogOutput("OnRecvOpen for RUNWAY request", "grey");

                    _simConnection.AddToFacilityDefinition(sd, "OPEN AIRPORT");
                    _simConnection.AddToFacilityDefinition(sd, "OPEN RUNWAY");

                    _simConnection.AddToFacilityDefinition(sd, "HEADING");
                    _simConnection.AddToFacilityDefinition(sd, "LENGTH");
                    _simConnection.AddToFacilityDefinition(sd, "PRIMARY_ILS_ICAO");
                    _simConnection.AddToFacilityDefinition(sd, "PRIMARY_ILS_REGION");
                    _simConnection.AddToFacilityDefinition(sd, "PRIMARY_NUMBER");
                    _simConnection.AddToFacilityDefinition(sd, "PRIMARY_DESIGNATOR");
                    _simConnection.AddToFacilityDefinition(sd, "SECONDARY_ILS_ICAO");
                    _simConnection.AddToFacilityDefinition(sd, "SECONDARY_ILS_REGION");
                    _simConnection.AddToFacilityDefinition(sd, "SECONDARY_NUMBER");
                    _simConnection.AddToFacilityDefinition(sd, "SECONDARY_DESIGNATOR");


                    _simConnection.AddToFacilityDefinition(sd, "CLOSE RUNWAY");
                    _simConnection.AddToFacilityDefinition(sd, "CLOSE AIRPORT");

                    _simConnection.RegisterFacilityDataDefineStruct<runway>(SIMCONNECT_FACILITY_DATA_TYPE.RUNWAY);
                    _simConnection.RequestFacilityData(sd, rd, Utils.reqICAO, "");

                    break;

                case "VOR":

                    VoiceAttackPlugin.LogOutput("OnRecvOpen for VOR request", "grey");

                    _simConnection.AddToFacilityDefinition(sd, "OPEN VOR");


                    _simConnection.AddToFacilityDefinition(sd, "GS_ALTITUDE");
                    _simConnection.AddToFacilityDefinition(sd, "HAS_GLIDE_SLOPE");
                    _simConnection.AddToFacilityDefinition(sd, "FREQUENCY");
                    _simConnection.AddToFacilityDefinition(sd, "TYPE");
                    _simConnection.AddToFacilityDefinition(sd, "LOCALIZER");
                    _simConnection.AddToFacilityDefinition(sd, "GLIDE_SLOPE");
                    _simConnection.AddToFacilityDefinition(sd, "NAME");


                    _simConnection.AddToFacilityDefinition(sd, "CLOSE VOR");

                    _simConnection.RegisterFacilityDataDefineStruct<vor>(SIMCONNECT_FACILITY_DATA_TYPE.VOR);
                    _simConnection.RequestFacilityData(sd, rd, Utils.reqVORICAO, Utils.reqVORregion);

                    break;

                case "SIMVAR":

                    VoiceAttackPlugin.LogOutput("OnRecvOpen for SimVar request", "grey");

                    try
                    {
                        _simConnection.AddToDataDefinition(DataDefinitions.SimVar, Utils.simvar, Utils.simvarunit, Utils.dataType, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                    }
                    catch (Exception e)
                    {
                        VoiceAttackPlugin.LogOutput("Failed to add data definition. " + e, "grey");
                    }

                    switch (Utils.resultDataType)
                    {

                        case "Int":

                            try
                            {

                                _simConnection.RegisterDataDefineStruct<SimVarInt>(DataDefinitions.SimVar);

                            }
                            catch (Exception e)
                            {
                                VoiceAttackPlugin.LogOutput("Failed to register struct definition. " + e, "grey");
                            }
                            break;

                        case "String":

                            try
                            {

                                _simConnection.RegisterDataDefineStruct<SimVarString>(DataDefinitions.SimVar);

                            }
                            catch (Exception e)
                            {
                                VoiceAttackPlugin.LogOutput("Failed to register struct definition. " + e, "grey");
                            }
                            break;

                        case "Double":

                            try
                            {

                                _simConnection.RegisterDataDefineStruct<SimVarDouble>(DataDefinitions.SimVar);

                            }
                            catch (Exception e)
                            {
                                VoiceAttackPlugin.LogOutput("Failed to register struct definition. " + e, "grey");
                            }
                            break;

                        case "Float":

                            try
                            {

                                _simConnection.RegisterDataDefineStruct<SimVarFloat>(DataDefinitions.SimVar);

                            }
                            catch (Exception e)
                            {
                                VoiceAttackPlugin.LogOutput("Failed to register struct definition. " + e, "grey");
                            }
                            break;

                        case "Bool":

                            try
                            {

                                _simConnection.RegisterDataDefineStruct<SimVarBool>(DataDefinitions.SimVar);

                            }
                            catch (Exception e)
                            {
                                VoiceAttackPlugin.LogOutput("Failed to register struct definition. " + e, "grey");
                            }
                            break;
                    }
                    
                    try
                    {
                        _simConnection.RequestDataOnSimObjectType(RequestTypes.SimVar, DataDefinitions.SimVar, 0, SIMCONNECT_SIMOBJECT_TYPE.USER);

                        requestPending[RequestTypes.SimVar] = true;

                        
                    }
                    catch (Exception e)
                    {
                        VoiceAttackPlugin.LogOutput("Failed to request data. " + e, "grey");
                    }

                    break;



            }


        }

        private void simconnect_Facility_OnRecvOpen(SimConnect sender, SIMCONNECT_RECV_OPEN data)
        {

            VoiceAttackPlugin.LogOutput(System.Reflection.MethodBase.GetCurrentMethod()?.Name, "grey");



            Connected = true;

        }

        private void simconnect_Facility_OnRecvQuit(SimConnect sender, SIMCONNECT_RECV data)
        {

            VoiceAttackPlugin.LogOutput(System.Reflection.MethodBase.GetCurrentMethod()?.Name, "grey");
            Disconnect();
        
        }

        private void simconnect_Facility_OnRecvFacilityData(SimConnect sender, SIMCONNECT_RECV_FACILITY_DATA data)
        {

            VoiceAttackPlugin.LogOutput(System.Reflection.MethodBase.GetCurrentMethod()?.Name, "grey");


            switch (requestType)
            {
                case "AIRPORT":

                    SIMCONNECT_FACILITY_DATA_TYPE t = (SIMCONNECT_FACILITY_DATA_TYPE)data.Type;

                    switch (t)
                    {

                        case SIMCONNECT_FACILITY_DATA_TYPE.AIRPORT:

                            try
                            {
                                airport a = (airport)data.Data[0];

                                VoiceAttackPlugin.LogOutput("latitude: " + a.latitude, "grey");
                                VoiceAttackPlugin.LogOutput("arrivals: " + a.arrivals, "grey");
                                VoiceAttackPlugin.LogOutput("nRunways: " + a.nRunways, "grey");

                                isLastMessageReceived = true;

                            }
                            catch (Exception ex)
                            {
                                VoiceAttackPlugin.LogErrorOutput("OnRecvFacilityData problem with AIRPORT type. " + ex, "red");
                            }

                            break;

                        case SIMCONNECT_FACILITY_DATA_TYPE.RUNWAY:


                            break;

                    }

                    break;

                case "RUNWAY":

                    SIMCONNECT_FACILITY_DATA_TYPE s = (SIMCONNECT_FACILITY_DATA_TYPE)data.Type;

                    switch (s)
                    {

                        case SIMCONNECT_FACILITY_DATA_TYPE.AIRPORT:

                            try
                            {

                            }
                            catch (Exception ex)
                            {
                                VoiceAttackPlugin.LogErrorOutput("OnRecvFacilityData problem with AIRPORT type. " + ex, "red");
                            }

                            break;

                        case SIMCONNECT_FACILITY_DATA_TYPE.RUNWAY:

                            try
                            {

                                runway a = (runway)data.Data[0];


                                //VoiceAttackPlugin.LogOutput("Runway Heading: " + a.heading, "grey");
                                //VoiceAttackPlugin.LogOutput("Runway Length: " + a.length, "grey");
                                VoiceAttackPlugin.LogOutput("Runway Primary Number: " + a.primaryNumber, "grey");
                                VoiceAttackPlugin.LogOutput("Runway Primary Designator: " + a.primaryDesignator, "grey");
                                VoiceAttackPlugin.LogOutput("Runway Primary ILS ICAO: " + a.primaryVORICAO, "grey");
                                VoiceAttackPlugin.LogOutput("Runway Primary ILS region: " + a.primaryVORregion, "grey");
                                VoiceAttackPlugin.LogOutput("Runway Secondary Number: " + a.secondaryNumber, "grey");
                                VoiceAttackPlugin.LogOutput("Runway Secondary Designator: " + a.secondaryDesignator, "grey");
                                VoiceAttackPlugin.LogOutput("Runway Secondary ILS ICAO: " + a.secondaryVORICAO, "grey");
                                VoiceAttackPlugin.LogOutput("Runway Secondary ILS region: " + a.secondaryVORregion, "grey");

                                Utils.simConnectMSGLoop++;

                                switch (Utils.simConnectMSGLoop)
                                {

                                    case 2:

                                        Utils.scHeading1 = a.heading;
                                        Utils.scLength1 = a.length;
                                        Utils.scPrimNumb1 = a.primaryNumber;
                                        Utils.scPrimDesign1 = a.primaryDesignator;
                                        Utils.scSecNumb1 = a.secondaryNumber;
                                        Utils.scSecDesign1 = a.secondaryDesignator;

                                        VoiceAttackPlugin.LogOutput("PrimNum: " + Utils.scPrimNumb1, "grey");
                                        VoiceAttackPlugin.LogOutput("SecNum: " + Utils.scSecNumb1, "grey");


                                        if (a.primaryVORICAO != "")
                                        {
                                            Utils.scPrimVORICAO1 = a.primaryVORICAO;
                                            Utils.scPrimVORregion1 = a.primaryVORregion;
                                        }

                                        if (a.secondaryVORICAO != "")
                                        {
                                            Utils.scSecVORICAO1 = a.secondaryVORICAO;
                                            Utils.scSecVORregion1 = a.secondaryVORregion;
                                        }

                                        break;

                                    case 3:

                                        Utils.scHeading2 = a.heading;
                                        Utils.scLength2 = a.length;
                                        Utils.scPrimNumb2 = a.primaryNumber;
                                        Utils.scPrimDesign2 = a.primaryDesignator;
                                        Utils.scSecNumb2 = a.secondaryNumber;
                                        Utils.scSecDesign2 = a.secondaryDesignator;

                                        VoiceAttackPlugin.LogOutput("PrimNum: " + Utils.scPrimNumb2, "grey");
                                        VoiceAttackPlugin.LogOutput("SecNum: " + Utils.scSecNumb2, "grey");


                                        if (a.primaryVORICAO != "")
                                        {
                                            Utils.scPrimVORICAO2 = a.primaryVORICAO;
                                            Utils.scPrimVORregion2 = a.primaryVORregion;
                                        }

                                        if (a.secondaryVORICAO != "")
                                        {
                                            Utils.scSecVORICAO2 = a.secondaryVORICAO;
                                            Utils.scSecVORregion2 = a.secondaryVORregion;
                                        }

                                        break;

                                    case 4:

                                        Utils.scHeading3 = a.heading;
                                        Utils.scLength3 = a.length;
                                        Utils.scPrimNumb3 = a.primaryNumber;
                                        Utils.scPrimDesign3 = a.primaryDesignator;
                                        Utils.scSecNumb3 = a.secondaryNumber;
                                        Utils.scSecDesign3 = a.secondaryDesignator;

                                        VoiceAttackPlugin.LogOutput("PrimNum: " + Utils.scPrimNumb3, "grey");
                                        VoiceAttackPlugin.LogOutput("SecNum: " + Utils.scSecNumb3, "grey");

                                        if (a.primaryVORICAO != "")
                                        {
                                            Utils.scPrimVORICAO3 = a.primaryVORICAO;
                                            Utils.scPrimVORregion3 = a.primaryVORregion;
                                        }

                                        if (a.secondaryVORICAO != "")
                                        {
                                            Utils.scSecVORICAO3 = a.secondaryVORICAO;
                                            Utils.scSecVORregion3 = a.secondaryVORregion;
                                        }

                                        break;

                                    case 5:

                                        Utils.scHeading4 = a.heading;
                                        Utils.scLength4 = a.length;
                                        Utils.scPrimNumb4 = a.primaryNumber;
                                        Utils.scPrimDesign4 = a.primaryDesignator;
                                        Utils.scSecNumb4 = a.secondaryNumber;
                                        Utils.scSecDesign4 = a.secondaryDesignator;

                                        VoiceAttackPlugin.LogOutput("PrimNum: " + Utils.scPrimNumb4, "grey");
                                        VoiceAttackPlugin.LogOutput("SecNum: " + Utils.scSecNumb4, "grey");

                                        if (a.primaryVORICAO != "")
                                        {
                                            Utils.scPrimVORICAO4 = a.primaryVORICAO;
                                            Utils.scPrimVORregion4 = a.primaryVORregion;
                                        }

                                        if (a.secondaryVORICAO != "")
                                        {
                                            Utils.scSecVORICAO4 = a.secondaryVORICAO;
                                            Utils.scSecVORregion4 = a.secondaryVORregion;
                                        }

                                        break;

                                    case 6:

                                        Utils.scHeading5 = a.heading;
                                        Utils.scLength5 = a.length;
                                        Utils.scPrimNumb5 = a.primaryNumber;
                                        Utils.scPrimDesign5 = a.primaryDesignator;
                                        Utils.scSecNumb5 = a.secondaryNumber;
                                        Utils.scSecDesign5 = a.secondaryDesignator;

                                        VoiceAttackPlugin.LogOutput("PrimNum: " + Utils.scPrimNumb5, "grey");
                                        VoiceAttackPlugin.LogOutput("SecNum: " + Utils.scSecNumb5, "grey");

                                        if (a.primaryVORICAO != "")
                                        {
                                            Utils.scPrimVORICAO5 = a.primaryVORICAO;
                                            Utils.scPrimVORregion5 = a.primaryVORregion;
                                        }

                                        if (a.secondaryVORICAO != "")
                                        {
                                            Utils.scSecVORICAO5 = a.secondaryVORICAO;
                                            Utils.scSecVORregion5 = a.secondaryVORregion;
                                        }

                                        break;

                                    case 7:

                                        Utils.scHeading6 = a.heading;
                                        Utils.scLength6 = a.length;
                                        Utils.scPrimNumb6 = a.primaryNumber;
                                        Utils.scPrimDesign6 = a.primaryDesignator;
                                        Utils.scSecNumb6 = a.secondaryNumber;
                                        Utils.scSecDesign6 = a.secondaryDesignator;

                                        VoiceAttackPlugin.LogOutput("PrimNum: " + Utils.scPrimNumb6, "grey");
                                        VoiceAttackPlugin.LogOutput("SecNum: " + Utils.scSecNumb6, "grey");

                                        if (a.primaryVORICAO != "")
                                        {
                                            Utils.scPrimVORICAO6 = a.primaryVORICAO;
                                            Utils.scPrimVORregion6 = a.primaryVORregion;
                                        }

                                        if (a.secondaryVORICAO != "")
                                        {
                                            Utils.scSecVORICAO6 = a.secondaryVORICAO;
                                            Utils.scSecVORregion6 = a.secondaryVORregion;
                                        }

                                        break;

                                    case 8:

                                        Utils.scHeading7 = a.heading;
                                        Utils.scLength7 = a.length;
                                        Utils.scPrimNumb7 = a.primaryNumber;
                                        Utils.scPrimDesign7 = a.primaryDesignator;
                                        Utils.scSecNumb7 = a.secondaryNumber;
                                        Utils.scSecDesign7 = a.secondaryDesignator;

                                        VoiceAttackPlugin.LogOutput("PrimNum: " + Utils.scPrimNumb7, "grey");
                                        VoiceAttackPlugin.LogOutput("SecNum: " + Utils.scSecNumb7, "grey");

                                        if (a.primaryVORICAO != "")
                                        {
                                            Utils.scPrimVORICAO7 = a.primaryVORICAO;
                                            Utils.scPrimVORregion7 = a.primaryVORregion;
                                        }

                                        if (a.secondaryVORICAO != "")
                                        {
                                            Utils.scSecVORICAO7 = a.secondaryVORICAO;
                                            Utils.scSecVORregion7 = a.secondaryVORregion;
                                        }

                                        break;

                                    case 9:

                                        Utils.scHeading8 = a.heading;
                                        Utils.scLength8 = a.length;
                                        Utils.scPrimNumb8 = a.primaryNumber;
                                        Utils.scPrimDesign8 = a.primaryDesignator;
                                        Utils.scSecNumb8 = a.secondaryNumber;
                                        Utils.scSecDesign8 = a.secondaryDesignator;

                                        VoiceAttackPlugin.LogOutput("PrimNum: " + Utils.scPrimNumb8, "grey");
                                        VoiceAttackPlugin.LogOutput("SecNum: " + Utils.scSecNumb8, "grey");

                                        if (a.primaryVORICAO != "")
                                        {
                                            Utils.scPrimVORICAO8 = a.primaryVORICAO;
                                            Utils.scPrimVORregion8 = a.primaryVORregion;
                                        }

                                        if (a.secondaryVORICAO != "")
                                        {
                                            Utils.scSecVORICAO8 = a.secondaryVORICAO;
                                            Utils.scSecVORregion8 = a.secondaryVORregion;
                                        }

                                        break;

                                }

                                isLastMessageReceived = true;

                            }
                            catch (Exception ex)
                            {
                                VoiceAttackPlugin.LogErrorOutput("OnRecvFacilityData problem with RUNWAY type. " + ex, "red");
                            }

                            break;



                    }

                    break;

                case "VOR":

                    SIMCONNECT_FACILITY_DATA_TYPE r = (SIMCONNECT_FACILITY_DATA_TYPE)data.Type;

                    switch (r)
                    {

                        case SIMCONNECT_FACILITY_DATA_TYPE.VOR:

                            try
                            {


                                vor a = (vor)data.Data[0];

                                VoiceAttackPlugin.LogOutput("Glide Slope Altitude: " + a.gsAltitude, "grey");
                                VoiceAttackPlugin.LogOutput("Has Glide Slope: " + a.hasGlideSlope, "grey");
                                float freq = a.frequency;
                                freq = freq / 1000000;
                                VoiceAttackPlugin.LogOutput("Frequency: " + freq, "grey");
                                VoiceAttackPlugin.LogOutput("Type: " + a.type, "grey");
                                VoiceAttackPlugin.LogOutput("Localizer Heading: " + a.localizer, "grey");
                                VoiceAttackPlugin.LogOutput("Glide Slope Angle: " + a.glide_slope, "grey");
                                VoiceAttackPlugin.LogOutput("Name: " + a.name, "grey");

                                Utils.scLocFreq = freq;
                                Utils.scLocHeading = a.localizer;
                                Utils.scLocName = a.name;

                                VoiceAttackPlugin.LogOutput("Name: " + Utils.scLocName, "grey");

                                isLastMessageReceived = true;

                            }
                            catch (Exception ex)
                            {
                                VoiceAttackPlugin.LogErrorOutput("OnRecvFacilityData problem with VOR type. " + ex, "red");
                            }

                            break;

                    }

                    break;

            }
                     
        }

        private void simconnect_Facility_OnRecvFacilityDataEnd(SimConnect sender, SIMCONNECT_RECV_FACILITY_DATA_END data)
        {
            VoiceAttackPlugin.LogOutput(System.Reflection.MethodBase.GetCurrentMethod()?.Name, "grey");
        }

        /// <summary>
        /// Fires when the sim is closed/exited.  
        /// </summary>
        /// 

        private void CheckMessagesTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            _simConnection?.ReceiveMessage();
        }



        private void simconnect_OnRecvQuit(SimConnect sender, SIMCONNECT_RECV data)
        {

            VoiceAttackPlugin.LogOutput(System.Reflection.MethodBase.GetCurrentMethod()?.Name, "grey");

        }

        /// <summary>
        /// Fires when new data is recieved from the sim
        /// </summary>
        private void simconnect_OnRecvSimobjectDataBytype(SimConnect sender, SIMCONNECT_RECV_SIMOBJECT_DATA_BYTYPE data)

        {
            VoiceAttackPlugin.LogOutput(System.Reflection.MethodBase.GetCurrentMethod()?.Name, "grey");

            switch ((DataDefinitions)data.dwDefineID)
            {

                case DataDefinitions.SimVar:
                    {

                        switch (Utils.resultDataType)
                        {
                            case "String":

                                try
                                {                                 

                                    GetSimVarS = (SimVarString)data.dwData[0];

                                    Utils.resultDataString = GetSimVarS.simvar;

                                    isLastMessageReceived = true;

                                }
                                catch (Exception e)

                                {
                                    VoiceAttackPlugin.LogOutput("Problem casting String. " + e, "grey");
                                }

                                VoiceAttackPlugin.LogOutput("SimVar String: " + Utils.resultDataString, "grey");

                                break;

                            case "Double":

                                try
                                {                                    

                                    GetSimVarD = (SimVarDouble)data.dwData[0];

                                    Utils.resultDataValue = GetSimVarD.simvar;

                                    isLastMessageReceived = true;

                                }
                                catch (Exception e)

                                {
                                    VoiceAttackPlugin.LogOutput("Problem casting Double. " + e, "grey");
                                }

                                VoiceAttackPlugin.LogOutput("SimVar Double: " + Utils.resultDataValue, "grey");

                                break;

                            case "Int":

                                try
                                {                                   

                                    GetSimVarI = (SimVarInt)data.dwData[0];

                                    Utils.resultDataValue = Convert.ToDouble(GetSimVarI.simvar);

                                    isLastMessageReceived = true;

                                }
                                catch (Exception e)

                                {
                                    VoiceAttackPlugin.LogOutput("Problem casting Int. " + e, "grey");
                                }

                                VoiceAttackPlugin.LogOutput("SimVar Int: " + Utils.resultDataValue, "grey");

                                break;

                            case "Bool":

                                try
                                {                                    

                                    GetSimVarB = (SimVarBool)data.dwData[0];

                                    Utils.resultDataValue = Convert.ToDouble(GetSimVarB.simvar);

                                    isLastMessageReceived = true;

                                }
                                catch (Exception e)

                                {
                                    VoiceAttackPlugin.LogOutput("Problem casting Bool. " + e, "grey");
                                }


                                VoiceAttackPlugin.LogOutput("SimVar Bool: " + Utils.resultDataValue, "grey");

                                break;

                            case "Float":

                                try
                                {                                    

                                    GetSimVarF = (SimVarFloat)data.dwData[0];

                                    Utils.resultDataValue = Convert.ToDouble(GetSimVarF.simvar);

                                    isLastMessageReceived = true;

                                }
                                catch (Exception e)

                                {
                                    VoiceAttackPlugin.LogOutput("Problem casting float. " + e, "grey");
                                }


                                VoiceAttackPlugin.LogOutput("SimVar Float: " + Utils.resultDataValue, "grey");

                                break;

                            default:
                                VoiceAttackPlugin.LogOutput("SimVar not getting any type", "grey");
                                break;

                        }

                        break;
                    }

                default:
                    VoiceAttackPlugin.LogOutput("data.dwDefineID false", "grey");
                    break;
            }
            

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
                catch (Exception e)
                {
                    VoiceAttackPlugin.LogErrorOutput("MessageReceiveThreadHandler failed. " + e.Message, "red");
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
            VoiceAttackPlugin.LogOutput("Checked for messages.", "grey");
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
            catch (Exception e)
            {
                VoiceAttackPlugin.LogErrorOutput("Failed to initiate event handlers. " + e.Message, "red");
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

        public void DefineRequestType(string type)
        {

            requestType = type;

        }



    }

}
