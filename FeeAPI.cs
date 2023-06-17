// Author: Yifan Wang

using System.ComponentModel;

using FS.API;

namespace AMLConversion_ClassLibrary
{
    public class FeeAPI : ObservableObject
    {
        ///////////////////////////////////////////////////////////////////////////////////
        //
        // Constructor and attributes
        //
        ///////////////////////////////////////////////////////////////////////////////////
#region Constructor and attributes

        // Global variables
        public CoreApi Api { get; set; }

        private string _ipAddress;
        private int _port;
        private string _username;
        private string _password;
        private bool _isApiConnected;
        private string _connectionState;
        private string _simulationState;

        public FeeAPI(string ipAdress = "localhost", int port = 0xFEE, string username = "admin", string password = "admin")
        {
            Api = new CoreApi();
            Api.ApiNetwork.PropertyChanged += NetworkOnPropertyChanged;

            _ipAddress = ipAdress;
            _port = port;
            _username = username;
            _password = password;
            _simulationState = "Standby";
        }

        public string IPAddress
        {
            get => _ipAddress;
            set
            {
                _ipAddress = value;
                OnPropertyChanged("IPAddress");
            }
        }

        public int Port
        {
            get => _port;
            set
            {
                _port = value;
                OnPropertyChanged("Port");
            }
        }

        public string Username
        {
            get => _username;
            set
            {
                _username = value;
                OnPropertyChanged("Username");
            }
        }

        public string Password
        {
            get => _password;
            set
            {
                _password = value;
                OnPropertyChanged("Password");
            }
        }

        public bool IsApiConnected
        {
            get => _isApiConnected;
            set
            {
                _isApiConnected = value;
                OnPropertyChanged("IsApiConnected");
            }
        }

        public string ConnectionState
        {
            get => _connectionState;
            set
            {
                _connectionState = value;
                OnPropertyChanged("ConnectionState");
            }
        }

        public string SimulationState
        {
            get => _simulationState;
            set
            {
                _connectionState = value;
                OnPropertyChanged("SimulationState");
            }
        }

#endregion

        ///////////////////////////////////////////////////////////////////////////////////
        //
        // PropertyChanged event listener
        //
        ///////////////////////////////////////////////////////////////////////////////////
#region PropertyChanged event listener

        private void NetworkOnPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(ApiNetwork.State))
            {
                ConnectionState = ((ApiNetwork)sender).State.ToString();

                if (((ApiNetwork)sender).State.ToString() == "Connected")
                {
                    IsApiConnected = true;
                }
                else
                {
                    IsApiConnected = false;
                }
            }
        }

#endregion

        ///////////////////////////////////////////////////////////////////////////////////
        //
        // Private functions
        //
        ///////////////////////////////////////////////////////////////////////////////////
#region Private functions

        private ApiVariableDefinition Get_interfaceVariable(Guid interfaceGUID, string variableName)
        {
            ApiVariableDefinition returnValue = null;

            var variables = Api.Interface.GetVariablesFromInterface(interfaceGUID.ToString());
            foreach (var variable in variables)
            {
                if (variable.Tag == variableName)
                    returnValue = variable;
            }

            return returnValue;
        }

        private Guid Get_interfaceGUID(string name)
        {
            var interfaceGUIDs = Api.Interface.GetAllInterfaces();
            var returnGUID = Guid.Empty;

            foreach (var interfaceGUID in interfaceGUIDs)
            {
                var interfaceProperties = Api.Interface.GetInterfaceProperties(Guid.Parse(interfaceGUID));
                foreach (var property in interfaceProperties)
                {
                    if (property.PropertyName == "InterfaceName" && property.PropertyValue == name)
                        returnGUID = Guid.Parse(interfaceGUID);
                }
            }

            return returnGUID;
        }

#endregion

        ///////////////////////////////////////////////////////////////////////////////////
        //
        // Public functions
        //
        ///////////////////////////////////////////////////////////////////////////////////
#region Public functions

        public void Connect()
        {
            if (!IsApiConnected)
            {
                //Port 4078 is always the same
                Api.Connect(IPAddress, Port, Username, Password);
            }
        }

        public void Disconnect(bool deleteEventHandler = false)
        {
            ConnectionState = "Disconnecting";

            //the API needs to be disconnected if it will not be used anymore
            if(deleteEventHandler)
                Api.ApiNetwork.PropertyChanged -= NetworkOnPropertyChanged;
            Api.Disconnect();

            ConnectionState = "Disconnected";
            IsApiConnected = false;
        }

        public void Process_command(string command, string commandVariableTag)
        {
            var interfaceGUID = Get_interfaceGUID("Symbolic Variables");
            var commandStateVariableTag = commandVariableTag + "State";

            // Set symbolic variable value in fee
            SimulationState = " -> [" + commandVariableTag + "] " + command;
            var commandVariable = Get_interfaceVariable(interfaceGUID, commandVariableTag);
            Api.Interface.ForceVariableByGuid(commandVariable.VariableGuid, command);

            // Check for command finish
            var commandState = "undefined";
            while (!commandState.Contains(command + "f"))
            {
                var commandStateVariable = Get_interfaceVariable(interfaceGUID, commandStateVariableTag);  // refresh variable
                commandState = commandStateVariable.Value.ToString();
                if (!commandState.Contains(command + "f"))
                    Thread.Sleep(10);  // read interval
            }
            SimulationState = " <- [" + commandVariableTag + "] " + commandState;
            Api.Interface.RemoveForceVariable(commandVariable.VariableGuid);

            SimulationState = "Standby";
        }

#endregion
    }
}
