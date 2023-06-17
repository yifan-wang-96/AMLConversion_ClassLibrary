// Author: Yifan Wang

using System.IO.Ports;
using System.ComponentModel;

using Aml.Engine.CAEX;
using Aml.Engine.AmlObjects;

namespace AMLConversion_ClassLibrary
{
    public class Simulator: ObservableObject
    {
        ///////////////////////////////////////////////////////////////////////////////////
        //
        // Constructor
        //
        ///////////////////////////////////////////////////////////////////////////////////
#region Constructor and attributes

        // Global variables
        private StatusInformation _status;
        private bool _isScanned;
        private FeeAPI _api;
        private SerialCommunication _serial;

        public Simulator(FeeAPI api = null, SerialCommunication ser = null)  // Constructor
        {
            _status = new StatusInformation();
            _status.PropertyChanged += StatusOnPropertyChanged;
            _api = api;
            _serial = ser;
            _isScanned = false;
        }

        public FeeAPI Api
        {
            get => _api;
            set
            {
                _api = value;
            }
        }

        public SerialCommunication Serial
        {
            get => _serial;
            set
            {
                _serial = value;
            }
        }

        public bool IsScanned
        {
            get => _isScanned;
            set
            {
                _isScanned = value;
                OnPropertyChanged("IsScanned");
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
            else if (e.PropertyName == "IsWorking")
                OnPropertyChanged("StatusIsWorking");
        }

#endregion

        ///////////////////////////////////////////////////////////////////////////////////
        //
        // Private functions
        //
        ///////////////////////////////////////////////////////////////////////////////////
#region Private functions

        private List<string[]> Get_commands(string filePath)
        {
            Status.Text = "Gathering Information";

            var commandList = new List<string[]>();

            // Create Indexers
            var plantfileDocument = CAEXDocument.LoadFromFile(filePath);
            var IH = plantfileDocument.CAEXFile.InstanceHierarchy["DemonstratorProject"];
            var amlBRCL = AutomationMLBaseRoleClassLibType.RoleClassLib(plantfileDocument);
            var RC = amlBRCL.ProcessStructure;
            var processesIE = AMLCustom.Get_IEs(IH, RC).First();

            var RCL = plantfileDocument.CAEXFile.RoleClassLib["ProcessRoleClassLib"];
            RC = RCL.RoleClass["Job"];
            var jobIEs = AMLCustom.Get_IEs(processesIE, RC);

            foreach (var jobIE in jobIEs)
            {
                RC = RCL.RoleClass["ProductionProcess"];
                var productionProcessIEs = AMLCustom.Get_IEs(jobIE, RC);

                foreach (var productionProcessIE in productionProcessIEs)
                {
                    var productionStepIEs = productionProcessIE.InternalElement;

                    foreach (var productionStepIE in productionStepIEs)
                    {
                        var productionStepRC = RCL.RoleClass["ProductionStep"];

                        if (productionStepIE.HasRoleClassReference(productionStepRC.RoleClass["Reference"]))
                        {
                            commandList.Add(new[] { "reference", "SSArduino" });
                        }
                        else if (productionStepIE.HasRoleClassReference(productionStepRC.RoleClass["Home"]))
                        {
                            commandList.Add(new[] { "home", "SSArduino" });
                        }
                        else if (productionStepIE.HasRoleClassReference(productionStepRC.RoleClass["Park"]))
                        {
                            commandList.Add(new[] { "park", "SSArduino" });
                        }
                        else if (productionStepIE.HasRoleClassReference(productionStepRC.RoleClass["MoveTo"]))
                        {
                            var id = productionStepIE.Attribute["slotID"].Value;
                            commandList.Add(new[] { "moveTo_SLOT" + id, "SSArduino" });
                        }
                        else if (productionStepIE.HasRoleClassReference(productionStepRC.RoleClass["MoveToSide"]))
                        {
                            var id = productionStepIE.Attribute["slotID"].Value;
                            commandList.Add(new[] { "moveToSide_SLOT" + id, "ARArduino" });
                        }
                        else if (productionStepIE.HasRoleClassReference(productionStepRC.RoleClass["LEDCheck"]))
                        {
                            commandList.Add(new[] { "ledCheck", "ARArduino" });
                        }
                        else if (productionStepIE.HasRoleClassReference(productionStepRC.RoleClass["TakeProductAsGripper"]))
                        {
                            var color = productionStepIE.Attribute["color"].Value;
                            commandList.Add(new[] { "takeProductAsGripper_COLOR" + color, "ARArduino" });
                        }
                        else if (productionStepIE.HasRoleClassReference(productionStepRC.RoleClass["DropProductAsGripper"]))
                        {
                            commandList.Add(new[] { "dropProductAsGripper", "ARArduino" });
                        }
                        else if (productionStepIE.HasRoleClassReference(productionStepRC.RoleClass["TakeProductAsStation"]))
                        {
                            var id = productionStepIE.Attribute["slotID"].Value;
                            var color = productionStepIE.Attribute["color"].Value;
                            commandList.Add(new[] { "takeProductAsStation_SLOT" + id + "_COLOR" + color, "SSArduino" });
                        }
                        else if (productionStepIE.HasRoleClassReference(productionStepRC.RoleClass["DropProductAsStation"]))
                        {
                            var id = productionStepIE.Attribute["slotID"].Value;
                            commandList.Add(new[] { "dropProductAsStation_SLOT" + id, "SSArduino" });
                        }
                    }
                }
            }

            return commandList;
        }
        
        private void Simulate_VC(string[] command)
        {
            // Create Indexers
            
            string commandVariableTag = null;
            if (command[1] == "ARArduino")
            {
                commandVariableTag = "ar_command";
            }
            else if (command[1] == "SSArduino")
            {
                commandVariableTag = "ss_command";
            }

            Api.Process_command(command[0], commandVariableTag);
        }

        private void Simulate_PD(string[] command)
        {
            Tuple<string, SerialPort> serialCommand = Tuple.Create("unknown serial port", new SerialPort());

            if(command[1] == "ARArduino")
                serialCommand = Tuple.Create(command[0], Serial.ARPort);
            else if (command[1] == "SSArduino")
                serialCommand = Tuple.Create(command[0], Serial.SSPort);

            Serial.Process_command(serialCommand);
        }

        private void Choose_simulation(string name, string[] command)
        {
            if (name == "PD")
                Simulate_PD(command);
            else if (name == "VC")
                Simulate_VC(command);
        }

#endregion

        ///////////////////////////////////////////////////////////////////////////////////
        //
        // Public functions
        //
        ///////////////////////////////////////////////////////////////////////////////////
#region Public functions
        public async Task<List<string[]>> Scan_slotPropertiesAsync()
        {
            return await Task.Run(() => Scan_slotProperties());
        }

        public List<string[]> Scan_slotProperties()
        {
            Status.Text = "start scan";
            Status.Progress = 0;

            // Create command Lists
            var commandList = new List<Tuple<string, SerialPort>>();
            var initCommands = new List<Tuple<string, SerialPort>>
            {
                Tuple.Create("moveToSide_SLOT0", Serial.ARPort),
                Tuple.Create("reference", Serial.SSPort),
                Tuple.Create("home", Serial.SSPort),
                Tuple.Create("ledCheck", Serial.ARPort)
            };
            var checkSlotsCommands = new List<Tuple<string, SerialPort>>();
            for (var i = 1; i <= 10; i++)
            {
                checkSlotsCommands.Add(Tuple.Create("moveTo_SLOT" + i.ToString(), Serial.SSPort));
                checkSlotsCommands.Add(Tuple.Create("readProperties_SLOT" + i.ToString(), Serial.ARPort));
            }

            // Merge all commandLists into one commandList
            foreach (var command in initCommands)
                commandList.Add(command);
            foreach (var command in checkSlotsCommands)
                commandList.Add(command);

            // Process merged commandList
            var slotProperties = new List<string[]>
            {
                new []{"1", "", ""},
                new []{"2", "", ""},
                new []{"3", "", ""},
                new []{"4", "", ""},
                new []{"5", "", ""},
                new []{"6", "", ""},
                new []{"7", "", ""},
                new []{"8", "", ""},
                new []{"9", "", ""},
                new []{"10", "", ""}
            };

            Status.Set_taskCount(commandList.Count);
            foreach (var command in commandList)
            {
                Status.Text = "process command: " + command.Item1;

                var responseArray = Serial.Process_command(command);

                if (responseArray.Length > 0)  // if serialConnection sends back slotProperties information
                {
                    for (var i = 0; i < 10; i++)
                    {
                        if (slotProperties[i][0] == responseArray[1])
                        {
                            slotProperties[i][1] = responseArray[2];
                            slotProperties[i][2] = responseArray[3];
                        }
                    }
                }

                Status.Finish_task();
            }

            IsScanned = true;
            Status.Text = "scan finished";

            return slotProperties;
        }

        public async Task SimulateAsync(string filePath, bool simulatePhysicalDemonstrator, bool simulateVirtualCommissioning)
        {
            Status.Text = "start simulation";

            // Gathering information [Task 1]
            var commandList = Get_commands(filePath);
            Status.Set_taskCount(commandList.Count);

            foreach (var command in commandList)
            {
                Status.Text = "process command: " + command[0];

                var taskList = new List<string>();
                if (simulateVirtualCommissioning && !simulatePhysicalDemonstrator)
                    taskList.Add("VC");
                if (simulatePhysicalDemonstrator && !simulateVirtualCommissioning)
                    taskList.Add("PD");
                else if (simulateVirtualCommissioning && simulatePhysicalDemonstrator)
                {
                    while (!(Serial.SSConnectionState == "Standby" && Serial.ARConnectionState == "Standby" && Api.SimulationState == "Standby"))
                    {
                        // wait
                    }

                    taskList.Add("PD");
                    taskList.Add("VC");
                }

                // Parallel asynchronous task execution
                await Task.Run(() =>
                {
                    Parallel.ForEach<string>(taskList, (task) =>
                    {
                        Choose_simulation(task, command);
                    });
                });

                Status.Finish_task();
            }
            
            Status.Text = "simulation finished";
        }

#endregion

    }
}
