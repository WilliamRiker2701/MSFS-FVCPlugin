//=================================================================================================================
// PROJECT: MSFS Agent
// PURPOSE: This class does the interfacing to SimConnect for Microsoft Flight Simulator 2020
// AUTHOR: James Clark
// Licensed under the MS-PL license. See LICENSE.md file in the project root for full license information.
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


namespace MSFS
{
    /// Main class that operates against the flight simulator
    public class Agent
    {
        Microsoft.FlightSimulator.SimConnect.SimConnect _simConnection;
        WASimCommander.CLI.Client.WASimClient _waSimConnection;

        const int WM_USER_SIMCONNECT = 0x402;
        bool _connected = false;
        bool _wasmconnected = false;


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
                    _waSimConnection.connectServer();
                    WAServerConnected = true;

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

                _simConnection.UnsubscribeFromSystemEvent(EventTypes.SIMSTART);
                _simConnection.UnsubscribeFromSystemEvent(EventTypes.SIMSTOP);
                _simConnection.UnsubscribeFromSystemEvent(EventTypes.PAUSE);

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

        /// <summary>
        /// Initializes all the events we want to be able to use with the sim.
        /// The events are automatically extracted from the "EventTypes" enum.
        /// </summary>
        private void initEvents()
        {

            // maps an event for each entry found in the EventTypes enum
            foreach (EventTypes item in Enum.GetValues(typeof(EventTypes)))
            {
                _simConnection.MapClientEventToSimEvent(item, EventFactory.GetEventName(item));
                _simConnection.AddClientEventToNotificationGroup(NOTIFICATION_GROUPS.DEFAULT, item, false);
                Debug.WriteLine("Event mapped: ", EventFactory.GetEventName(item));
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
        /// Adds the data definitions for each type of data request
        /// </summary>
        public void AddDataDefinitions()
        {
            try
            {
                // The order of fields must match the exact order in the structure for the data definition, the data types also be compatible

                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "AIRSPEED INDICATED", "Knots", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "AMBIENT TEMPERATURE", "Fahrenheit", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "APU SWITCH", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "APU GENERATOR SWITCH", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "APU PCT RPM", "Percent Over 100", SIMCONNECT_DATATYPE.FLOAT64, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "ATC AIRLINE", null, SIMCONNECT_DATATYPE.STRING256, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "ATC FLIGHT NUMBER", null, SIMCONNECT_DATATYPE.STRING256, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "ATC ID", null, SIMCONNECT_DATATYPE.STRING256, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "AUTO BRAKE SWITCH CB", "Number", SIMCONNECT_DATATYPE.FLOAT64, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "AUTOPILOT AIRSPEED HOLD", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "AUTOPILOT AIRSPEED HOLD VAR", "Knots", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "AUTOPILOT ALTITUDE LOCK", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "AUTOPILOT ALTITUDE LOCK VAR", "Feet", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "AUTOPILOT APPROACH HOLD", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "AUTOPILOT ATTITUDE HOLD", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "AUTOPILOT AVAILABLE", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "AUTOPILOT BACKCOURSE HOLD", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "AUTOPILOT DISENGAGED", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "AUTOPILOT FLIGHT DIRECTOR ACTIVE", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "AUTOPILOT HEADING LOCK", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "AUTOPILOT HEADING LOCK DIR", "Degrees", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "AUTOPILOT MASTER", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "AUTOPILOT MAX BANK ID", null, SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "AUTOPILOT NAV SELECTED", "Number", SIMCONNECT_DATATYPE.FLOAT64, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "AUTOPILOT NAV1 LOCK", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "AUTOPILOT THROTTLE ARM", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "AUTOPILOT VERTICAL HOLD", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "AUTOPILOT VERTICAL HOLD VAR", "Feet/minute", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "AUTOPILOT YAW DAMPER", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "BLEED AIR SOURCE CONTROL", "Enum", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "BRAKE PARKING INDICATOR", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "CIRCUIT SWITCH ON:17", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "CIRCUIT SWITCH ON:18", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "CIRCUIT SWITCH ON:19", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "CIRCUIT SWITCH ON:20", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "CIRCUIT SWITCH ON:21", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "CIRCUIT SWITCH ON:22", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "COM ACTIVE FREQUENCY:1", "MHz", SIMCONNECT_DATATYPE.FLOAT64, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "COM STANDBY FREQUENCY:1", "MHz", SIMCONNECT_DATATYPE.FLOAT64, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "COM ACTIVE FREQUENCY:2", "MHz", SIMCONNECT_DATATYPE.FLOAT64, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "COM STANDBY FREQUENCY:2", "MHz", SIMCONNECT_DATATYPE.FLOAT64, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "ELECTRICAL MASTER BATTERY", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "EXTERNAL POWER ON", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "ENG ANTI ICE:1", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "ENG ANTI ICE:2", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "ENG N1 RPM:1", null, SIMCONNECT_DATATYPE.FLOAT64, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "ENG N1 RPM:2", null, SIMCONNECT_DATATYPE.FLOAT64, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "ENG N2 RPM:1", null, SIMCONNECT_DATATYPE.FLOAT64, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "ENG N2 RPM:2", null, SIMCONNECT_DATATYPE.FLOAT64, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "ENGINE TYPE", "Enum", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "FLAPS HANDLE INDEX", "Number", SIMCONNECT_DATATYPE.FLOAT64, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "FLAPS HANDLE PERCENT", "Percent Over 100", SIMCONNECT_DATATYPE.FLOAT64, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "FUELSYSTEM PUMP SWITCH:1", null, SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "FUELSYSTEM PUMP SWITCH:2", null, SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "FUELSYSTEM PUMP SWITCH:3", null, SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "FUELSYSTEM PUMP SWITCH:4", null, SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "FUELSYSTEM PUMP SWITCH:5", null, SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "FUELSYSTEM PUMP SWITCH:6", null, SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "FUELSYSTEM VALVE SWITCH:1", null, SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "FUELSYSTEM VALVE SWITCH:2", null, SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "FUELSYSTEM VALVE SWITCH:3", null, SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "GEAR HANDLE POSITION", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "GENERAL ENG STARTER:1", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "GENERAL ENG STARTER:2", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "GROUND VELOCITY", "Knots", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "HEADING INDICATOR", "Degrees", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "HYDRAULIC SWITCH", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "IS GEAR RETRACTABLE", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "LIGHT BEACON", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "LIGHT CABIN", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "LIGHT LANDING", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "LIGHT LOGO", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "LIGHT NAV", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "LIGHT PANEL", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "LIGHT RECOGNITION", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "LIGHT STROBE", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "LIGHT TAXI", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "LIGHT WING", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "LOCAL TIME", "Hours", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "MASTER IGNITION SWITCH", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "NAV ACTIVE FREQUENCY:1", "MHz", SIMCONNECT_DATATYPE.FLOAT64, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "NAV STANDBY FREQUENCY:1", "MHz", SIMCONNECT_DATATYPE.FLOAT64, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "NAV ACTIVE FREQUENCY:2", "MHz", SIMCONNECT_DATATYPE.FLOAT64, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "NAV STANDBY FREQUENCY:2", "MHz", SIMCONNECT_DATATYPE.FLOAT64, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "NUMBER OF ENGINES", "Number", SIMCONNECT_DATATYPE.FLOAT64, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "PANEL ANTI ICE SWITCH", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "PITOT HEAT", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "PLANE ALT ABOVE GROUND", "Feet", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "PLANE ALTITUDE", "Feet", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "PLANE LATITUDE", "Degrees", SIMCONNECT_DATATYPE.FLOAT64, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "PLANE LONGITUDE", "Degrees", SIMCONNECT_DATATYPE.FLOAT64, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "PROP DEICE SWITCH:1", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "PUSHBACK STATE", "Enum", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "SIM ON GROUND", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "SPOILER AVAILABLE", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "SPOILERS ARMED", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "SPOILERS HANDLE POSITION", "Percent Over 100", SIMCONNECT_DATATYPE.FLOAT64, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "STRUCTURAL DEICE SWITCH", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "TITLE", null, SIMCONNECT_DATATYPE.STRING256, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "TRANSPONDER AVAILABLE", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "TRANSPONDER CODE:1", "Hz", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "VERTICAL SPEED", "Feet/minute", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "WATER RUDDER HANDLE POSITION", "Position", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                _simConnection.AddToDataDefinition(DataDefinitions.PlaneState, "WINDSHIELD DEICE SWITCH", "Bool", SIMCONNECT_DATATYPE.INT32, 0.0f, SimConnect.SIMCONNECT_UNUSED);

                // Each data definition must also have its struct registered otherwise we only get a unit value returned vs. the data we want
                _simConnection.RegisterDataDefineStruct<PlaneState>(DataDefinitions.PlaneState);

                Debug.WriteLine("Data definitions added...");

            }
            catch (Exception ex)
            {
                Debug.WriteLine("Failed to add data definitions. Msg:" + ex.Message);
            }
        }

        /// <summary>
        /// Submits a request into the sim for data.  
        /// </summary>
        /// <param name="requestType">The type of data request to make</param>
        /// <param name="dataDefinition">The data definition to use for the response</param>
        public void RequestData(RequestTypes requestType, DataDefinitions dataDefinition)
        {
            try
            {
                _simConnection.RequestDataOnSimObjectType(requestType, dataDefinition, 0, SIMCONNECT_SIMOBJECT_TYPE.USER);

                requestPending[requestType] = true;

                Debug.WriteLine("Request sent...");

            }
            catch (Exception ex)
            {
                Debug.WriteLine("Failed to initiate data request. Msg: " + ex.Message);
            }
        }

        /// <summary>
        /// Triggers an event in the sim
        /// </summary>
        /// <param name="eventType">The type of event to trigger</param>
        /// <param name="data">(Optional) data to pass with the event</param>
        public void TriggerEvent(EventTypes eventType, string data = "0")
        {

            // final data we will send to the sim with the event
            UInt32 eventData;
            Decimal d;

            if (String.IsNullOrWhiteSpace(data)) data = "0";

            try
            {
                switch (eventType)
                {

                    // all of these will fall through to be converted to BCD Hz values
                    case EventTypes.ADF_COMPLETE_SET:
                    case EventTypes.ADF2_COMPLETE_SET:
                    case EventTypes.COM_RADIO_SET:
                    case EventTypes.COM_STBY_RADIO_SET:
                    case EventTypes.COM2_RADIO_SET:
                    case EventTypes.COM2_STBY_RADIO_SET:
                    case EventTypes.NAV1_RADIO_SET:
                    case EventTypes.NAV1_STBY_SET:
                    case EventTypes.NAV2_RADIO_SET:
                    case EventTypes.NAV2_STBY_SET:

                        // frequency is expect to be in decimal form (i.e. 100.00)
                        if (Decimal.TryParse(data, out d))
                            eventData = Utils.Dec2Bcd(Decimal.ToUInt32(d * 100));
                        else
                            return;

                        break;

                    case EventTypes.XPNDR_SET:

                        // frequency is expect to be in NON-decimal form (i.e. 1000)
                        if (Decimal.TryParse(data, out d))
                            eventData = Utils.Dec2Bcd(Decimal.ToUInt32(d));
                        else
                            return;

                        break;

                    // for all other cases the data is converted to UInt32
                    default:

                        Byte[] Bytes = BitConverter.GetBytes(Convert.ToInt32(data));
                        eventData = BitConverter.ToUInt32(Bytes, 0);

                        break;

                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Failed to convert event data. EventData:{0}, Msg: {1}.", data, ex.Message);
                return;
            }

            _simConnection.TransmitClientEvent(SimConnect.SIMCONNECT_OBJECT_ID_USER, eventType, eventData, NOTIFICATION_GROUPS.DEFAULT, SIMCONNECT_EVENT_FLAG.GROUPID_IS_PRIORITY);

            Debug.WriteLine("Event sent...");

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

            string calcCode = dData + " (>L:" + varName + ")";



            _waSimConnection.executeCalculatorCode(calcCode, 0, out double fResult, out string sResult);

            Debug.WriteLine($"Calculator code '{calcCode}' returned: {fResult} and '{sResult}'", "<<");

            //----------------------------------------------------------------------------------

            //_waSimConnection.setVariable(new VariableRequest(varName), dData);

            //_waSimConnection.setLocalVariable(varName, dData);

            Debug.WriteLine("setLocalVariable() variable sent...");


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


#endregion


#region Event Handlers

/// <summary>
/// Initializes all the event handles for the agent
/// </summary>
private void initEventHandlers()
        {
            try
            {
                _simConnection.OnRecvOpen += new SimConnect.RecvOpenEventHandler(simconnect_OnRecvOpen);
                _simConnection.OnRecvQuit += new SimConnect.RecvQuitEventHandler(simconnect_OnRecvQuit);
                _simConnection.OnRecvException += new SimConnect.RecvExceptionEventHandler(simconnect_OnRecvException);
                _simConnection.OnRecvEvent += new SimConnect.RecvEventEventHandler(simconnect_OnRecvEvent);
                _simConnection.OnRecvSimobjectDataBytype += new SimConnect.RecvSimobjectDataBytypeEventHandler(simconnect_OnRecvSimobjectDataBytype);

                _simConnection.SubscribeToSystemEvent(EventTypes.SIMSTART, "SimStart");
                _simConnection.SubscribeToSystemEvent(EventTypes.SIMSTOP, "SimStop");

                // I can't recall why this is needed
                _simConnection.SetNotificationGroupPriority(NOTIFICATION_GROUPS.DEFAULT, SimConnect.SIMCONNECT_GROUP_PRIORITY_HIGHEST);

            }
            catch (Exception ex)
            {
                Debug.WriteLine("Failed to initiate event handlers. Msg:" + ex.Message);
            }
        }

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

            switch ((DataDefinitions)data.dwDefineID)
            {
                case DataDefinitions.PlaneState:
                    {
                        requestPending[RequestTypes.PlaneState] = false;

                        GetPlaneState = (PlaneState)data.dwData[0];

                        Debug.WriteLine("Airspeed_Indicated: " + GetPlaneState.Airspeed_Indicated.ToString());
                        Debug.WriteLine("Ambient_Temperature: " + GetPlaneState.Ambient_Temperature.ToString());
                        Debug.WriteLine("Apu_Switch: " + GetPlaneState.Apu_Switch.ToString());
                        Debug.WriteLine("Apu_Generator_Switch: " + GetPlaneState.Apu_Generator_Switch.ToString());
                        Debug.WriteLine("Apu_Pct_Rpm: " + GetPlaneState.Apu_Pct_Rpm.ToString());
                        Debug.WriteLine("Atc_Airline: " + GetPlaneState.Atc_Airline.ToString());
                        Debug.WriteLine("Atc_Flight_Number: " + GetPlaneState.Atc_Flight_Number.ToString());
                        Debug.WriteLine("Atc_Id: " + GetPlaneState.Atc_Id.ToString());
                        Debug.WriteLine("Auto_Brake_Switch_CB: " + GetPlaneState.Auto_Brake_Switch_CB.ToString());
                        Debug.WriteLine("Autopilot_Airspeed_Hold: " + GetPlaneState.Autopilot_Airspeed_Hold.ToString());
                        Debug.WriteLine("Autopilot_Airspeed_Hold_Var: " + GetPlaneState.Autopilot_Airspeed_Hold_Var.ToString());
                        Debug.WriteLine("Autopilot_Altitude_Lock: " + GetPlaneState.Autopilot_Altitude_Lock.ToString());
                        Debug.WriteLine("Autopilot_Altitude_Lock_Var: " + GetPlaneState.Autopilot_Altitude_Lock_Var.ToString());
                        Debug.WriteLine("Autopilot_Approach_Hold: " + GetPlaneState.Autopilot_Approach_Hold.ToString());
                        Debug.WriteLine("Autopilot_Attitude_Hold: " + GetPlaneState.Autopilot_Attitude_Hold.ToString());
                        Debug.WriteLine("Autopilot_Available: " + GetPlaneState.Autopilot_Available.ToString());
                        Debug.WriteLine("Autopilot_Backcourse_Hold: " + GetPlaneState.Autopilot_Backcourse_Hold.ToString());
                        Debug.WriteLine("Autopilot_Disengaged: " + GetPlaneState.Autopilot_Disengaged.ToString());
                        Debug.WriteLine("Autopilot_Flight_Director_Active: " + GetPlaneState.Autopilot_Flight_Director_Active.ToString());
                        Debug.WriteLine("Autopilot_Heading_Lock: " + GetPlaneState.Autopilot_Heading_Lock.ToString());
                        Debug.WriteLine("Autopilot_Heading_Lock_Dir: " + GetPlaneState.Autopilot_Heading_Lock_Dir.ToString());
                        Debug.WriteLine("Autopilot_Master: " + GetPlaneState.Autopilot_Master.ToString());
                        Debug.WriteLine("Autopilot_Max_Bank_ID: " + GetPlaneState.Autopilot_Max_Bank_ID.ToString());
                        Debug.WriteLine("Autopilot_Nav_Selected: " + GetPlaneState.Autopilot_Nav_Selected.ToString());
                        Debug.WriteLine("Autopilot_Nav1_Lock: " + GetPlaneState.Autopilot_Nav1_Lock.ToString());
                        Debug.WriteLine("Autopilot_Throttle_Arm: " + GetPlaneState.Autopilot_Throttle_Arm.ToString());
                        Debug.WriteLine("Autopilot_Vertical_Hold: " + GetPlaneState.Autopilot_Vertical_Hold.ToString());
                        Debug.WriteLine("Autopilot_Vertical_Hold_Var: " + GetPlaneState.Autopilot_Vertical_Hold_Var.ToString());
                        Debug.WriteLine("Autopilot_Yaw_Damper: " + GetPlaneState.Autopilot_Yaw_Damper.ToString());
                        Debug.WriteLine("Bleed_Air_Source_Control: " + GetPlaneState.Bleed_Air_Source_Control.ToString());
                        Debug.WriteLine("Brake_Parking_Indicator: " + GetPlaneState.Brake_Parking_Indicator.ToString());
                        Debug.WriteLine("Circuit_Switch_On_17: " + GetPlaneState.Circuit_Switch_On_17.ToString());
                        Debug.WriteLine("Circuit_Switch_On_18: " + GetPlaneState.Circuit_Switch_On_18.ToString());
                        Debug.WriteLine("Circuit_Switch_On_19: " + GetPlaneState.Circuit_Switch_On_19.ToString());
                        Debug.WriteLine("Circuit_Switch_On_20: " + GetPlaneState.Circuit_Switch_On_20.ToString());
                        Debug.WriteLine("Circuit_Switch_On_21: " + GetPlaneState.Circuit_Switch_On_21.ToString());
                        Debug.WriteLine("Circuit_Switch_On_22: " + GetPlaneState.Circuit_Switch_On_22.ToString());
                        Debug.WriteLine("Com1_Active_Frequency: " + GetPlaneState.Com1_Active_Frequency.ToString());
                        Debug.WriteLine("Com1_Standby_Frequency: " + GetPlaneState.Com1_Standby_Frequency.ToString());
                        Debug.WriteLine("Com2_Active_Frequency: " + GetPlaneState.Com2_Active_Frequency.ToString());
                        Debug.WriteLine("Com2_Standby_Frequency: " + GetPlaneState.Com2_Standby_Frequency.ToString());
                        Debug.WriteLine("Electrical_Master_Battery: " + GetPlaneState.Electrical_Master_Battery.ToString());
                        Debug.WriteLine("External_Power_On: " + GetPlaneState.External_Power_On.ToString());
                        Debug.WriteLine("Engine_Anti_Ice:1: " + GetPlaneState.Engine_Anti_Ice_1.ToString());
                        Debug.WriteLine("Engine_Anti_Ice:2: " + GetPlaneState.Engine_Anti_Ice_2.ToString());
                        Debug.WriteLine("Engine_N1_RPM:1: " + GetPlaneState.Engine_N1_RPM_1.ToString());
                        Debug.WriteLine("Engine_N1_RPM:2: " + GetPlaneState.Engine_N1_RPM_2.ToString());
                        Debug.WriteLine("Engine_N2_RPM:1: " + GetPlaneState.Engine_N2_RPM_1.ToString());
                        Debug.WriteLine("Engine_N2_RPM:2: " + GetPlaneState.Engine_N2_RPM_2.ToString());
                        Debug.WriteLine("Engine_Type: " + GetPlaneState.Engine_Type.ToString());
                        Debug.WriteLine("Flaps_Handle_Index: " + GetPlaneState.Flaps_Handle_Index.ToString());
                        Debug.WriteLine("Flaps_Handle_Percent: " + GetPlaneState.Flaps_Handle_Percent.ToString());
                        Debug.WriteLine("Fuelsystem_Pump_Switch:1: " + GetPlaneState.Fuelsystem_Pump_Switch_1.ToString());
                        Debug.WriteLine("Fuelsystem_Pump_Switch:2: " + GetPlaneState.Fuelsystem_Pump_Switch_2.ToString());
                        Debug.WriteLine("Fuelsystem_Pump_Switch:3: " + GetPlaneState.Fuelsystem_Pump_Switch_3.ToString());
                        Debug.WriteLine("Fuelsystem_Pump_Switch:4: " + GetPlaneState.Fuelsystem_Pump_Switch_4.ToString());
                        Debug.WriteLine("Fuelsystem_Pump_Switch:5: " + GetPlaneState.Fuelsystem_Pump_Switch_5.ToString());
                        Debug.WriteLine("Fuelsystem_Pump_Switch:6: " + GetPlaneState.Fuelsystem_Pump_Switch_6.ToString());
                        Debug.WriteLine("Fuelsystem_Valve_Switch:1" + GetPlaneState.Fuelsystem_Valve_Switch_1.ToString());
                        Debug.WriteLine("Fuelsystem_Valve_Switch:2" + GetPlaneState.Fuelsystem_Valve_Switch_2.ToString());
                        Debug.WriteLine("Fuelsystem_Valve_Switch:3" + GetPlaneState.Fuelsystem_Valve_Switch_3.ToString());
                        Debug.WriteLine("Gear_Handle_Position: " + GetPlaneState.Gear_Handle_Position.ToString());
                        Debug.WriteLine("General_Eng_Starter:1: " + GetPlaneState.General_Eng_Starter_1.ToString());
                        Debug.WriteLine("General_Eng_Starter:2: " + GetPlaneState.General_Eng_Starter_2.ToString());
                        Debug.WriteLine("Ground_Velocity: " + GetPlaneState.Ground_Velocity.ToString());
                        Debug.WriteLine("Heading_Indicator: " + GetPlaneState.Heading_Indicator.ToString());
                        Debug.WriteLine("Hydraulic_Switch: " + GetPlaneState.Hydraulic_Switch.ToString());
                        Debug.WriteLine("Is_Gear_Retractable: " + GetPlaneState.Is_Gear_Retractable.ToString());
                        Debug.WriteLine("Light_Beacon: " + GetPlaneState.Light_Beacon.ToString());
                        Debug.WriteLine("Light_Cabin: " + GetPlaneState.Light_Cabin.ToString());
                        Debug.WriteLine("Light_Landing: " + GetPlaneState.Light_Landing.ToString());
                        Debug.WriteLine("Light_Logo: " + GetPlaneState.Light_Logo.ToString());
                        Debug.WriteLine("Light_Nav: " + GetPlaneState.Light_Nav.ToString());
                        Debug.WriteLine("Light_Panel: " + GetPlaneState.Light_Panel.ToString());
                        Debug.WriteLine("Light_Recognition: " + GetPlaneState.Light_Recognition.ToString());
                        Debug.WriteLine("Light_Strobe: " + GetPlaneState.Light_Strobe.ToString());
                        Debug.WriteLine("Light_Taxi: " + GetPlaneState.Light_Taxi.ToString());
                        Debug.WriteLine("Light_Wing: " + GetPlaneState.Light_Wing.ToString());
                        Debug.WriteLine("Local_Time: " + GetPlaneState.Local_Time.ToString());
                        Debug.WriteLine("Master_Ignition_Switch: " + GetPlaneState.Master_Ignition_Switch.ToString());
                        Debug.WriteLine("Nav1_Active_Frequency: " + GetPlaneState.Nav1_Active_Frequency.ToString());
                        Debug.WriteLine("Nav1_Standby_Frequency: " + GetPlaneState.Nav1_Standby_Frequency.ToString());
                        Debug.WriteLine("Nav2_Active_Frequency: " + GetPlaneState.Nav2_Active_Frequency.ToString());
                        Debug.WriteLine("Nav2_Standby_Frequency: " + GetPlaneState.Nav2_Standby_Frequency.ToString());
                        Debug.WriteLine("Number_Of_Engines: " + GetPlaneState.Number_Of_Engines.ToString());
                        Debug.WriteLine("Panel_Anti_Ice_Switch: " + GetPlaneState.Panel_Anti_Ice_Switch.ToString());
                        Debug.WriteLine("Pitot_Heat: " + GetPlaneState.Pitot_Heat.ToString());
                        Debug.WriteLine("Plane_Alt_Above_Ground: " + GetPlaneState.Plane_Alt_Above_Ground.ToString());
                        Debug.WriteLine("Plane_Altitude: " + GetPlaneState.Plane_Altitude.ToString());
                        Debug.WriteLine("Plane_Latitude: " + GetPlaneState.Plane_Latitude.ToString());
                        Debug.WriteLine("Plane_Longitude: " + GetPlaneState.Plane_Longitude.ToString());
                        Debug.WriteLine("Prop_Deice_Switch: " + GetPlaneState.Prop_Deice_Switch.ToString());
                        Debug.WriteLine("Pushback_State: " + GetPlaneState.Pushback_State.ToString());
                        Debug.WriteLine("Sim_On_Ground: " + GetPlaneState.Sim_On_Ground.ToString());
                        Debug.WriteLine("Spoiler_Available: " + GetPlaneState.Spoiler_Available.ToString());
                        Debug.WriteLine("Spoilers_Armed: " + GetPlaneState.Spoiler_Available.ToString());
                        Debug.WriteLine("Spoilers_Handle_Position: " + GetPlaneState.Spoilers_Handle_Position.ToString());
                        Debug.WriteLine("Structural_Deice_Switch: " + GetPlaneState.Structural_Deice_Switch.ToString());
                        Debug.WriteLine("Title: " + GetPlaneState.Title.ToString());
                        Debug.WriteLine("Transponder_Available: " + GetPlaneState.Transponder_Available.ToString());
                        Debug.WriteLine("Transponder_Code: " + GetPlaneState.Transponder_Code.ToString());
                        Debug.WriteLine("Vertical_Speed: " + GetPlaneState.Vertical_Speed.ToString());
                        Debug.WriteLine("Water_Rudder_Handle_Position: " + GetPlaneState.Water_Rudder_Handle_Position.ToString());
                        Debug.WriteLine("Windshield_Deice_Switch: " + GetPlaneState.Windshield_Deice_Switch.ToString());

                        break;
                    }


                default:
                    break;

            }
        }

        /// <summary>
        /// Fires when a registered system event happens inside the sim (i.e. start, stop, pause)
        /// </summary>
        private void simconnect_OnRecvEvent(SimConnect sender, SIMCONNECT_RECV_EVENT recEvent)
        {

            // I'm still try to figure out when these get triggered.  
            // They often fire rapidly when exiting and entering flights.            

            switch ((EventTypes)recEvent.uEventID)
            {
                case EventTypes.SIMSTART:

                    Debug.WriteLine("Sim running");
                    break;

                case EventTypes.SIMSTOP:

                    Debug.WriteLine("Sim stopped");
                    break;

                case EventTypes.PAUSE:

                    Debug.WriteLine("Sim paused");
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
