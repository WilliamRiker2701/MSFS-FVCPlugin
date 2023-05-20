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
                    Debug.WriteLine("Unable to connect to sim. Msg:" + ex.Message);
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
                    Debug.WriteLine("Unable to connect to WASM. Msg:" + ex.Message);
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
                    Debug.WriteLine("Unable to connect to WASM. Msg:" + ex.Message);
                }
            }



        }

        public void WASMConnect2()
        {

            if (_waSimConnection == null)
            {

                Debug.WriteLine("WASM client not initialized");

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

                Debug.WriteLine("WASM Client could not connect to SimConnect for unknown reason");

            if (hr != HR.OK)
            {
                WasmStatus = WasmModuleStatus.NotFound;
                Debug.WriteLine("WASM Server not found or couldn't connect");
                return;
            }

            WasmStatus = WasmModuleStatus.Connected;
            WAServerConnected = _waSimConnection.isConnected();
            Debug.WriteLine("Connected to WASimConnect Server");                    


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
                    Debug.WriteLine("Unable to connect to WASM. Msg:" + ex.Message);
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

                Debug.WriteLine("Connection to sim closed");
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Failed to Disconnect and clean up. Msg:" + ex.Message);
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

                Debug.WriteLine("Connection to WASM closed");
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Failed to Disconnect WASM. Msg:" + ex.Message);
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
            double dData;

            if (String.IsNullOrWhiteSpace(varData)) varData = "0";

            try
            {

                dData = Convert.ToDouble(varData);

            }
            catch (Exception ex)
            {
                Debug.WriteLine("Failed to trigger WASM method. Data:{0}, Msg: {1}.", varData, ex.Message);
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

        public double TriggerCalcCode(string calcCode = "0")
        {


            //------CALCULATOR CODE--------------------------------------------------------------



            _waSimConnection.executeCalculatorCode(calcCode, CalcResultType.Double, out double fResult, out string sResult);

            Debug.WriteLine($"Calculator code '{calcCode}' returned: {fResult} and '{sResult}'", "<<");

            //----------------------------------------------------------------------------------

            //_waSimConnection.setVariable(new VariableRequest(varName), dData);

            //_waSimConnection.setLocalVariable(varName, dData);

            Debug.WriteLine("setLocalVariable() variable sent...");

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
                Debug.WriteLine("Failed to send event data. Data:{0}, Msg: {1}.", varData, ex.Message);
                return;
            }


        }

        public string TriggerReqPMDG(string varName)
        {

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
                Debug.WriteLine("Failed to receive variable data. Variable:{0}, Msg: {1}.", varName, ex.Message);
                return "NULL";
            }


        }
        public void TriggerComKey(string varName, string varData = "0")
        {

            //------CALCULATOR CODE--------------------------------------------------------------

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

            Debug.WriteLine($"Calculator code '{calcCode}' returned: {fResult} and '{sResult}'", "<<");

            //----------------------------------------------------------------------------------

            //_waSimConnection.setVariable(new VariableRequest(varName), dData);

            //_waSimConnection.setLocalVariable(varName, dData);

            Debug.WriteLine("setLocalVariable() variable sent...");


        }

        public void TriggerKey(string varName, string varData = "0")
        {

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
                Debug.WriteLine("Failed to send event data. Data:{0}, Msg: {1}.", varData, ex.Message);
                return;
            }


        }

        public void TriggerKeySimconnect(FsControlList varName, string varData = "0")
        {

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
                Debug.WriteLine("Failed to convert data. Key Data:{0}, Msg: {1}.", varData, ex.Message);
                return;
            }

            _simConnection.TransmitClientEvent(SimConnect.SIMCONNECT_OBJECT_ID_USER, varName, kData, NOTIFICATION_GROUPS.DEFAULT, SIMCONNECT_EVENT_FLAG.GROUPID_IS_PRIORITY);

            Debug.WriteLine("Event sent...");

        }

        public double TriggerReqSimVar(string varName)
        {
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

                        Debug.WriteLine("Executed calculator code with 2 numbers: " + calcCode);

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

                    Debug.WriteLine("Executed calculator code with 1 number: " + calcCode);

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

                Debug.WriteLine("Executed calculator code: " + calcCode);

                if (varName == "TRANSPONDER CODE")
                {

                    fResult = fResult * 1000000;

                }

                return fResult;


            }


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

            Console.WriteLine("SimConnect_OnRecvException: " + ex.ToString());


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
            Debug.WriteLine("SimConnect_OnRecvOpen");

            Connected = true;

        }

        /// <summary>
        /// Fires when the sim is closed/exited.  
        /// </summary>
        private void simconnect_OnRecvQuit(SimConnect sender, SIMCONNECT_RECV data)
        {
            Debug.WriteLine("Sim has exited.");
            Disconnect();
        }

        /// <summary>
        /// Fires when new data is recieved from the sim
        /// </summary>
        private void simconnect_OnRecvSimobjectDataBytype(SimConnect sender, SIMCONNECT_RECV_SIMOBJECT_DATA_BYTYPE data)

        {
            Debug.WriteLine("SimConnect_OnRecvSimobjectDataBytype");


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
                    Debug.WriteLine("MessageReceiveThreadHandler failed. Msg:" + ex.Message);
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
            Debug.WriteLine("Checked for messages (manually)");
        }

        static void ClientStatusHandler(ClientEvent ev)
        {
            Debug.WriteLine($"Client event {ev.eventType} - \"{ev.message}\"; Client status: {ev.status}", "^^");
        }

        static void LogHandler(LogRecord lr, LogSource src)
        {
            Debug.WriteLine($"{src} Log: {lr}", "@@");  // LogRecord has a convenience ToString() override
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
                Debug.WriteLine("Failed to initiate event handlers. Msg:" + ex.Message);
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
            Console.Write($"<< Got Data for request {(Requests)dr.requestId} \"{dr.nameOrCode}\" with Value: ");
            // Convert the received data into a value using DataRequestRecord's tryConvert() methods.
            // This could be more efficient in a "real" application, but it's good enough for our tests with only 2 value types.
            if (dr.tryConvert(out float fVal))
                Console.WriteLine($"(float) {fVal}");
            else if (dr.tryConvert(out string sVal))
            {
                Console.WriteLine($"(string) \"{sVal}\"");
            }
            else
                Console.WriteLine("Could not convert result data to value!");

        }

    }
}
