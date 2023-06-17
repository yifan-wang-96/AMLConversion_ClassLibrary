// Author: Yifan Wang

using System.IO.Ports;
using System.Management;
using System.ComponentModel;

namespace AMLConversion_ClassLibrary
{
    public class SerialCommunication : ObservableObject
    {
        ///////////////////////////////////////////////////////////////////////////////////
        //
        // Constructor and attributes
        //
        ///////////////////////////////////////////////////////////////////////////////////

#region  Constructor and attributes

        private int _baudrate;
        private bool _isARConnected;
        private bool _isSSConnected;
        private string _arConnectionState;
        private string _ssConnectionState;
        private StatusInformation _status;
        private SerialPort _arPort;
        private SerialPort _ssPort;
        private List<SerialPort> _arduinoPorts = new List<SerialPort> ();

        public SerialCommunication(int baudrate = 115200)  // Constructor
        {
            _status = new StatusInformation ();
            _status.PropertyChanged += StatusOnPropertyChanged;
            _baudrate = baudrate;
            ARConnectionState = "Disconnected";
            SSConnectionState = "Disconnected";
        }

        public SerialPort ARPort
        {
            get => _arPort;
            set
            {
                _arPort = value;
            }
        }

        public SerialPort SSPort
        {
            get => _ssPort;
            set
            {
                _ssPort = value;
            }
        }

        public bool IsARConnected
        {
            get => _isARConnected;
            set
            {
                if (value != _isARConnected)
                {
                    // Refresh StatusInformation
                    if (value == true && IsSSConnected)
                    {
                        Status.Text = "Connection established";
                    }
                    else if (value == false && !IsSSConnected)
                    {
                        Status.Text = "Disconnected";
                    }
                }

                _isARConnected = value;
                OnPropertyChanged("IsARConnected");
            }
        }

        public bool IsSSConnected
        {
            get => _isSSConnected;
            set
            {
                if(value != _isSSConnected)
                {
                    // Refresh StatusInformation
                    if (IsARConnected && value == true)
                    {
                        Status.Text = "Connection established";
                    }
                    else if (!IsARConnected && value == false)
                    {
                        Status.Text = "Disconnected";
                    }
                }
                
                _isSSConnected = value;
                OnPropertyChanged("IsSSConnected");
            }
        }

        public string ARConnectionState
        {
            get => _arConnectionState;
            set
            {
                if (value != _arConnectionState)
                {
                    // Refresh IsARConnected
                    if (value == "Processing" || value == "Standby")
                    {
                        IsARConnected = true;
                    }
                    else
                    {
                        IsARConnected = false;
                    }

                    _arConnectionState = value;
                    OnPropertyChanged("ARConnectionState");
                }
            }
        }

        public string SSConnectionState
        {
            get => _ssConnectionState;
            set
            {
                if(value != _ssConnectionState)
                {
                    _ssConnectionState = value;

                    // Refresh IsSSConnected
                    if (value == "Processing" || value == "Standby")
                    {
                        IsSSConnected = true;
                    }
                    else
                    {
                        IsSSConnected = false;
                    }

                    OnPropertyChanged("SSConnectionState");
                }  
            }
        }

        public StatusInformation Status
        {
            get => _status;
            set
            {
                _status = value;
            }
        }

        #endregion

        ///////////////////////////////////////////////////////////////////////////////////
        //
        // PropertyChanged event listener
        //
        ///////////////////////////////////////////////////////////////////////////////////
#region PropertyChanged event listener

        private void StatusOnPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "Text")
                OnPropertyChanged("StatusText");
            else if (e.PropertyName == "Progress")
                OnPropertyChanged("StatusProgress");
        }

#endregion

        ///////////////////////////////////////////////////////////////////////////////////
        //
        // Private functions
        //
        ///////////////////////////////////////////////////////////////////////////////////

#region Private functions

        private ManagementObject[] FindPorts()
        {
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_PnPEntity");
                List<ManagementObject> objects = new List<ManagementObject>();

                foreach(ManagementObject obj in searcher.Get())
                {
                    objects.Add(obj);
                }

                return objects.ToArray();
            }
            catch (Exception ex)
            {
                return new ManagementObject[] { };
            }
        }

        private string ParseCOMName(ManagementObject obj)
        {
            string name = obj["Name"].ToString();
            int startIndex = name.LastIndexOf("(");
            int endIndex = name.LastIndexOf(")");

            if (startIndex != -1 && endIndex != -1)
            {
                name = name.Substring(startIndex + 1, endIndex - startIndex - 1);
                return name;
            }

            return null;
        }

        private List<SerialPort> FindPortByManufacturerList(List<string> manufacturerList)
        {
            var comNames = new List<SerialPort>();

            foreach (ManagementObject obj in FindPorts())  // find all ports
            {
                if (obj["Manufacturer"] != null)  // if SerialPort object has a manufacturer attribute
                {
                    foreach (var manufacturer in manufacturerList)  // check for arduino manufacturers
                    {
                        if (obj["Manufacturer"].ToString().ToLower().Contains(manufacturer.ToLower()))
                        {
                            string comName = ParseCOMName(obj);  // get COM name

                            if (comName != null)
                            {
                                comNames.Add(new SerialPort(comName, _baudrate));
                            }
                        }
                    }
                }
            }
            return comNames;
        }

        private string[] Read_string(SerialPort port)
        {
            var response = port.ReadLine();
            var lastIndex = response.Count() - 1;
            response = response.Remove(lastIndex, 1);  // remove the last 2 characters

            if (response != "")
            {
                Status.Text = " <- [" + port.PortName + "] " + response;

                // handshake responses
                if (response.Contains("ar_arduino"))
                {
                    _arPort = port;
                    Write_string(port, "ok");
                    ARConnectionState = "Standby";
                    Status.Text = "Assignment finished: Port " + port.PortName;
                }
                else if (response.Contains("ss_arduino"))
                {
                    _ssPort = port;
                    Write_string(port, "ok");
                    SSConnectionState = "Standby";
                    Status.Text = "Assignment finished: Port " + port.PortName;
                }
                else  // main communication responses
                {
                    if (response.Contains("readPropertiesf"))
                    {
                        var responseFiltered = response.Split('_');

                        if (port.PortName == ARPort.PortName)
                            ARConnectionState = "Standby";
                        else if (port.PortName == SSPort.PortName)
                            SSConnectionState = "Standby";

                        return responseFiltered;
                    }
                    else if (response.Contains("f"))
                    {
                        if (port.PortName == ARPort.PortName)
                            ARConnectionState = "Standby";
                        else if (port.PortName == SSPort.PortName)
                            SSConnectionState = "Standby";
                    }
                }
            }

            return new string[] { };
        }

        private void Write_string(SerialPort port, string text)
        {
            if(text != "handshake" && text != "ok")
            {
                if (port.PortName == ARPort.PortName)
                    ARConnectionState = "Processing";
                else if (port.PortName == SSPort.PortName)
                    SSConnectionState = "Processing";
            }
            Status.Text = " -> [" + port.PortName + "] " + text;
            port.WriteLine(text);
        }

#endregion

        ///////////////////////////////////////////////////////////////////////////////////
        //
        // Public functions
        //
        ///////////////////////////////////////////////////////////////////////////////////

#region Public functions

        public void Disconnect()
        {
            if (_ssPort != null)
                _ssPort.Close();
            SSConnectionState = "Disconnected";
            if (_arPort != null)
                _arPort.Close();
            ARConnectionState = "Disconnected";
        }

        public void Connect()
        {
            Status.Text = "Search for Arduino ports";

            var manufacturers = new List<string> { "wch", "duino" };
            _arduinoPorts = FindPortByManufacturerList(manufacturers);

            if (_arduinoPorts.Count  == 0)
            {
                Status.Text = "No Arduino ports found";
            }
            else
            {
                foreach (var arduinoport in _arduinoPorts)
                {
                    if (!arduinoport.IsOpen)
                    {
                        Status.Text = "Open Port " + arduinoport.PortName;
                        arduinoport.Open();
                    }

                    Status.Text = "Assign Port " + arduinoport.PortName;

                    while (!Status.Text.Contains("Assignment finished"))
                    {
                        Write_string(arduinoport, "handshake");
                        Read_string(arduinoport);
                        Thread.Sleep(100);
                    }

                    arduinoport.DiscardInBuffer();
                    arduinoport.DiscardOutBuffer();
                }
            }
        }

        public string[] Process_command(Tuple<string, SerialPort> command)
        {
            Write_string(command.Item2, command.Item1);

            var responseArray = new string[] { };
            while (!(ARConnectionState == "Standby" && SSConnectionState == "Standby"))
            {
                responseArray = Read_string(command.Item2);
            }
            return responseArray;
        }

#endregion
    }
}
