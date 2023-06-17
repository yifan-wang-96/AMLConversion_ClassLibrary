// Author: Yifan Wang

using System.ComponentModel;

using Aml.Engine.CAEX;
using Aml.Engine.AmlObjects;

using FS.API;
using FS.SDK;
using FS.SDK.Io;
using FS.SDK.Extensibility.Contracts;
using Vector3 = FS.SDK.Mathematics.Vector3;

namespace AMLConversion_ClassLibrary
{
    public class Importer_fee : ObservableObject
    {
        ///////////////////////////////////////////////////////////////////////////////////
        //
        // Constructor and attributes
        //
        ///////////////////////////////////////////////////////////////////////////////////
#region  Constructor and attributes

        private StatusInformation _status;
        private FeeAPI _api;
        private List<string[]> _feeLibGuids;
        private List<Tuple<Guid, string, IOType, string>> _symbVars = new List<Tuple<Guid, string, IOType, string>> ();
        private Guid _symbVarInterfaceGUID = Guid.NewGuid(); //guid of the new interface instance

        public Importer_fee(FeeAPI api) // Constructor
        {
            _status = new StatusInformation();
            _status.PropertyChanged += StatusOnPropertyChanged;
            Api = api;
        }

        public StatusInformation Status
        {
            get => _status;
            set
            {
                _status = value;
            }
        }

        public FeeAPI Api
        {
            get => _api;
            set
            {
                _api = value;
            }
        }

        public List<string[]> FeeLibraryGuids
        {
            get => _feeLibGuids;
            set {
                _feeLibGuids = value;
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

        private void Set_defaultLibraryGuids()
        {
            Status.Text = "set default library guids";

            FeeLibraryGuids = new List<string[]>
            {
                new string[] {"frame", "ff85bbe5-2aeb-4c01-bb64-0f8dfc12098f"},
                new string[] {"transportUnit", "7dd7e406-719e-4c11-957a-45de93030239"},
                new string[] {"slot", "a0f5aa08-e603-464e-93f5-20d836e016f4"},
                new string[] {"productSink_red", "5b9b2b48-69ee-4904-a439-34b337179755"},
                new string[] {"productSink_green", "5b9b2b48-69ee-4904-a439-34b337179755"},
                new string[] {"productSink_blue", "5b9b2b48-69ee-4904-a439-34b337179755"},
                new string[] {"productSource_red", "5b9b2b48-69ee-4904-a439-34b337179755"},
                new string[] {"productSource_green", "5b9b2b48-69ee-4904-a439-34b337179755"},
                new string[] {"productSource_blue", "5b9b2b48-69ee-4904-a439-34b337179755"},
                new string[] {"Logic_ar_arduino", "0b555d63-9999-4fc4-b63b-146eee64d83a"},
                new string[] {"Logic_ss_arduino", "8e78b9ca-8222-4437-9e69-d27b6e6fe800" }
            };

            Status.Finish_task();
        }

        private async Task Import_libraryGuidsAsync(string APIDirectoryPath)
        {
            await Task.Run(() => Import_libraryGuids(APIDirectoryPath));
        }

        private void Import_libraryGuids(string APIDirectoryPath)
        {
            Status.Text = "import library guids";

            FeeLibraryGuids = new List<string[]>();
            var libraryObjectPaths = Directory.GetDirectories(APIDirectoryPath, "*", SearchOption.TopDirectoryOnly);
            
            foreach (var libraryObjectPath in libraryObjectPaths)
            {
                var libraryObjectName = Path.GetFileName(libraryObjectPath);

                var guid = "";
                using (var csvFile = new StreamReader(libraryObjectPath + @"\guid.txt"))
                {
                    var i = 0;
                    while (!csvFile.EndOfStream)
                    {
                        guid = csvFile.ReadLine();
                    }
                }

                var libraryGUID = new[] { libraryObjectName, guid };
                FeeLibraryGuids.Add(libraryGUID);
            }

            Status.Finish_task();
        }

        private async Task<string> Import_libraryObjectAsync(string libraryObjectName, string name = null,
            float xPos = 0, float yPos = 0, float zPos = 0,
            float xRot = 0, float yRot = 0, float zRot = 0)
        {
            Status.Text = "import " + libraryObjectName;

            var treeChildGUID = "";
            var libraryObjectGUID = "";

            // search for library object GUID
            var libSearch = FeeLibraryGuids.Where(element => element[0] == libraryObjectName);
            if (libSearch.Count() > 0)
            {
                libraryObjectGUID = FeeLibraryGuids.Where(element => element[0] == libraryObjectName).First()[1];

                if (libraryObjectGUID != null)
                {
                    // import library object
                    var treeLibObjGUID = (await Api.Api.Library.InsertTemplateElementLatestVersionAsync(libraryObjectGUID)).Item3;

                    // manipulate properties of library object
                    Api.Api.Object.CreateObject("LibraryObject", treeLibObjGUID);
                    Api.Api.Object.SetProperty(treeLibObjGUID, "LocalPosition", new Vector3(xPos, yPos, zPos), "Transform");
                    Api.Api.Object.SetProperty(treeLibObjGUID, "LocalRotation", new Vector3(xRot, yRot, zRot), "Transform");
                    await Api.Api.Object.SendAndWait(Guid.Parse(treeLibObjGUID));

                    // search for guid in child object
                    treeChildGUID = Api.Api.Object.GetAllChildrenFromSceneObject(treeLibObjGUID)[0];

                    // manipulate properties of child object
                    Api.Api.Object.CreateObject("CADAssembly", treeChildGUID);
                    if (name != null)
                        Api.Api.Object.SetProperty(treeChildGUID, nameof(SceneObject.Name), name);
                    Api.Api.Object.SetProperty(treeChildGUID, nameof(SceneObject.Name), name);
                    await Api.Api.Object.SendAndWait(Guid.Parse(treeChildGUID));
                }
                else
                {
                    Status.Text = "no Library Object in VC Software Fee found: " + libraryObjectName;
                }
            }
            else
            {
                Status.Text = "no entry in FeeLibraryGuids List found: " + libraryObjectName;
            }

            Status.Finish_task();

            if (treeChildGUID == "" || libraryObjectGUID == "")
                return null;
            else
                return treeChildGUID;
        }

        private async Task Remove_libraryRootElementsAsnyc()
        {
            Status.Text = "remove library elements";

            var libraryObjects = await Api.Api.Object.GetSceneObjectGuidsOfTypeAsync("LibraryObject");

            foreach (var libObj in libraryObjects)
            {
                var result = await Api.Api.Library.DecomposeLibraryElementAsync(libObj);
            }

            Status.Finish_task();
        }

        private async Task Create_diagramAsync(string diagramType, Guid parentGUID)  // overloaded function
        {
            await Create_diagramAsync(diagramType, new List<Guid>() { parentGUID }); 
        }

        private async Task Create_diagramAsync(string diagramType, List<Guid> parentGUIDs)
        {
            Status.Text = "create " + diagramType;

            // Create diagram
            var diagramGUID = await Api.Api.LogicAssigner.CreateDiagramOrFolder(diagramType, false, Guid.Empty) ?? throw new InvalidOperationException("Diagram was not created");

            var counter = 1;
            foreach(var parentGUID in parentGUIDs)
            {
                if (parentGUID != Guid.Empty)
                {
                    var inputVariables = new List<Tuple<Guid, string, IOType, string>>();
                    var inputLampVariables = new List<Tuple<Guid, string, IOType, string>>();
                    var outputVariables = new List<Tuple<Guid, string, IOType, string>>();
                    var outValVariables = new List<Tuple<Guid, string, IOType, string>>();
                    var createVariables = new List<Tuple<Guid, string, IOType, string>>();
                    var outValLogicObjectGUIDs = new List<Guid>();
                    var segmentedLampGUID = Guid.Empty;

                    // Create symbolic variable interface
                    var nameOfInterface = "Symbolic Variables";
                    await Api.Api.Interface.UpdateOrCreateInterfacePluginAsync(PluginGuids.Symbolicinterface, _symbVarInterfaceGUID, new Dictionary<string, object>  // Create a new interface Instance
                    {
                        { "InterfaceName", nameOfInterface }
                    });

                    // Create logic object slot container
                    if (await Get_treeGUIDWithNameAsync(parentGUID, "Logic") == "")
                        Status.Text = "no logic object found ... skip slot creation in diagram";
                    else
                    {
                        var logicObjectGUID = Guid.Empty;
                        if (diagramType == "kinematic_diagram")
                        {
                            // logicObjectGUID for kinematic_diagram is special
                            var sledgeGUID = Guid.Parse(await Get_treeGUIDWithNameAsync(parentGUID, "Sledge"));
                            logicObjectGUID = Guid.Parse(await Get_treeGUIDWithNameAsync(sledgeGUID, "Logic"));
                        }
                        else
                        {
                            logicObjectGUID = Guid.Parse(await Get_treeGUIDWithNameAsync(parentGUID, "Logic"));
                        }
                        await Api.Api.LogicAssigner.AddSlotContainer(diagramGUID, logicObjectGUID);  // Create logic object in diagram

                        switch (diagramType)
                        {
                            case "kinematic_diagram":
                                // Collect needed input connections
                                inputVariables = new List<Tuple<Guid, string, IOType, string>>()
                                {
                                    Tuple.Create(Guid.Empty, "sledge_pos", IOType.Real, "0"),
                                    Tuple.Create(Guid.Empty, "sledge_targetPos", IOType.Real, "0"),
                                    Tuple.Create(Guid.Empty, "sledge_defaultVel", IOType.Real, "0.1"),
                                    Tuple.Create(Guid.Empty, "sledge_posTol", IOType.Real, "0.004"),
                                    Tuple.Create(Guid.Empty, "arm_pos", IOType.Real, "0"),
                                    Tuple.Create(Guid.Empty, "arm_targetPos", IOType.Real, "0"),
                                    Tuple.Create(Guid.Empty, "arm_defaultVel", IOType.Real, "100"),
                                    Tuple.Create(Guid.Empty, "arm_posTol", IOType.Real, "1")
                                };

                                // Create motion joint slot containers
                                var linearachseMotionjointGUID = Guid.Parse(await Get_treeGUIDWithNameAsync(parentGUID, "Linearachse"));
                                await Api.Api.LogicAssigner.AddSlotContainer(diagramGUID, linearachseMotionjointGUID);
                                var schwenkachseMotionjointGUID = Guid.Parse(await Get_treeGUIDWithNameAsync(parentGUID, "Schwenkachse"));
                                await Api.Api.LogicAssigner.AddSlotContainer(diagramGUID, schwenkachseMotionjointGUID);
                                outValLogicObjectGUIDs = new List<Guid>() { linearachseMotionjointGUID, schwenkachseMotionjointGUID };

                                // Collect needed output connections
                                outValVariables = new List<Tuple<Guid, string, IOType, string>>()
                                {
                                    Tuple.Create(Guid.Empty, "sledge_pos", IOType.Real, "0"),
                                    Tuple.Create(Guid.Empty, "arm_pos", IOType.Real, "0")
                                };
                                break;

                            case "AR_arduino_diagram":
                                // Collect needed input connections
                                inputVariables = new List<Tuple<Guid, string, IOType, string>>()
                                {
                                    Tuple.Create(Guid.Empty, "ar_commandState", IOType.String, "undefined"),
                                    Tuple.Create(Guid.Empty, "ar_command", IOType.String, "standby"),
                                    Tuple.Create(Guid.Empty, "arm_pos", IOType.Real, "0"),
                                    Tuple.Create(Guid.Empty, "arm_posTol", IOType.Real, "1")
                                };

                                // Collect needed output connections
                                outputVariables = new List<Tuple<Guid, string, IOType, string>>()
                                {
                                    Tuple.Create(Guid.Empty, "arm_targetPos", IOType.Real, "0"),
                                    Tuple.Create(Guid.Empty, "ar_commandState", IOType.String, "undefined"),
                                    Tuple.Create(Guid.Empty, "arm_color", IOType.String, "undefined")
                                };
                                break;

                            case "SS_arduino_diagram":
                                // Collect needed input connections
                                inputVariables = new List<Tuple<Guid, string, IOType, string>>()
                                {
                                    Tuple.Create(Guid.Empty, "ss_commandState", IOType.String, "undefined"),
                                    Tuple.Create(Guid.Empty, "ss_command", IOType.String, "standby"),
                                    Tuple.Create(Guid.Empty, "sledge_pos", IOType.Real, "0"),
                                    Tuple.Create(Guid.Empty, "sledge_posTol", IOType.Real, "0.01")
                                };

                                // Collect needed output connections
                                outputVariables = new List<Tuple<Guid, string, IOType, string>>()
                                {
                                    Tuple.Create(Guid.Empty, "sledge_targetPos", IOType.Real, "0"),
                                    Tuple.Create(Guid.Empty, "sledge_defaultVel", IOType.Real, "0.1"),
                                    Tuple.Create(Guid.Empty, "ss_commandState", IOType.String, "undefined"),
                                    Tuple.Create(Guid.Empty, "slot1_color", IOType.String, "off"),
                                    Tuple.Create(Guid.Empty, "slot2_color", IOType.String, "off"),
                                    Tuple.Create(Guid.Empty, "slot3_color", IOType.String, "off"),
                                    Tuple.Create(Guid.Empty, "slot4_color", IOType.String, "off"),
                                    Tuple.Create(Guid.Empty, "slot5_color", IOType.String, "off"),
                                    Tuple.Create(Guid.Empty, "slot6_color", IOType.String, "off"),
                                    Tuple.Create(Guid.Empty, "slot7_color", IOType.String, "off"),
                                    Tuple.Create(Guid.Empty, "slot8_color", IOType.String, "off"),
                                    Tuple.Create(Guid.Empty, "slot9_color", IOType.String, "off"),
                                    Tuple.Create(Guid.Empty, "slot10_color", IOType.String, "off")
                                };
                                break;

                            case "stationLamps_diagram":
                                // Create segmented lamp slot containers
                                segmentedLampGUID = Guid.Parse(await Get_treeGUIDWithNameAsync(parentGUID, "Segmented lamp"));
                                await Api.Api.LogicAssigner.AddSlotContainer(diagramGUID, segmentedLampGUID);

                                // Collect needed station lamp input connections
                                inputLampVariables = new List<Tuple<Guid, string, IOType, string>>()
                                {
                                    Tuple.Create(Guid.Empty, "slot" + counter + "_color" , IOType.String, "undefined")
                                };
                                break;

                            case "armLamp_diagram":
                                // Create segmented lamp slot containers
                                segmentedLampGUID = Guid.Parse(await Get_treeGUIDWithNameAsync(parentGUID, "Segmented lamp"));
                                await Api.Api.LogicAssigner.AddSlotContainer(diagramGUID, segmentedLampGUID);

                                // Collect needed arm lamp input connections
                                inputLampVariables = new List<Tuple<Guid, string, IOType, string>>()
                                {
                                    Tuple.Create(Guid.Empty, "arm_color" , IOType.String, "undefined")
                                };
                                break;
                        }

                        if (inputVariables.Count > 0)
                        {
                            // Create and connect inputs
                            createVariables = await Delete_duplicatesAsync(inputVariables, _symbVars);
                            await Create_variablesAsync(_symbVarInterfaceGUID, createVariables);
                            inputVariables = await Copy_varGUIDAsync(_symbVars, inputVariables);
                            await Create_slotVarConnectionAsync("Input", new List<Guid>() { logicObjectGUID }, inputVariables);
                        }

                        if (inputLampVariables.Count > 0)
                        {
                            createVariables = await Delete_duplicatesAsync(inputLampVariables, _symbVars);
                            await Create_variablesAsync(_symbVarInterfaceGUID, createVariables);
                            inputLampVariables = await Copy_varGUIDAsync(_symbVars, inputLampVariables);
                            await Create_slotVarConnectionAsync("InputStation", new List<Guid>() { logicObjectGUID }, inputLampVariables);
                        }

                        if (outputVariables.Count > 0)
                        {
                            // Create and connect outputs
                            createVariables = await Delete_duplicatesAsync(outputVariables, _symbVars);
                            await Create_variablesAsync(_symbVarInterfaceGUID, createVariables);
                            outputVariables = await Copy_varGUIDAsync(_symbVars, outputVariables);
                            await Create_slotVarConnectionAsync("Output", new List<Guid>() { logicObjectGUID }, outputVariables);
                        }

                        if (outValVariables.Count > 0)
                        {
                            // Create and connect OutValue outputs
                            createVariables = await Delete_duplicatesAsync(outValVariables, _symbVars);
                            await Create_variablesAsync(_symbVarInterfaceGUID, createVariables);
                            outValVariables = await Copy_varGUIDAsync(_symbVars, outValVariables);
                            await Create_slotVarConnectionAsync("OutValue", outValLogicObjectGUIDs, outValVariables);
                        }
                    }
                }
                counter += 1;
            }
            Status.Finish_task();
        }

        private async Task<List<Tuple<Guid, string, IOType, string>>> Delete_duplicatesAsync(List<Tuple<Guid, string, IOType, string>> target, List<Tuple<Guid, string, IOType, string>> tool)
        {
            return await Task.Run(() => Delete_duplicates(target, tool));
        }

        private List<Tuple<Guid, string, IOType, string>> Delete_duplicates(List<Tuple<Guid, string, IOType, string>> target, List<Tuple<Guid, string, IOType, string>> tool)
        {
            var returnList = new List<Tuple<Guid, string, IOType, string>>();

            // check for already existing variables
            foreach (var targetVar in target)
            {
                var skipReturn = false;
                foreach (var toolVar in tool)
                {
                    if (toolVar.Item2 == targetVar.Item2)
                    {
                        skipReturn = true;
                    }
                }

                if (!skipReturn)
                    returnList.Add(targetVar);
            }

            return returnList;
        }

        private async Task Create_variablesAsync(Guid variableInterfaceGUID, List<Tuple<Guid, string, IOType, string>> variableList)
        {
            foreach (var variable in variableList)
            {
                var variableGUID = Guid.NewGuid();

                await Api.Api.Interface.UpdateOrCreateVariableAsync(PluginGuids.Symbolicinterface, variableInterfaceGUID, variableGUID, new ApiVariableDefinition
                {
                    Tag = variable.Item2,
                    Address = "",
                    Path = "",
                    Type = variable.Item3,
                    Comment = "",
                    Usage = IOMode.None,
                    Cycle = IOCycle.Continous
                }, variable.Item4); //value of the Variable

                var variableWithGUID = Tuple.Create(variableGUID, variable.Item2, variable.Item3, variable.Item4);
                _symbVars.Add(variableWithGUID);
            }
        }

        private async Task<List<Tuple<Guid, string, IOType, string>>> Copy_varGUIDAsync(List<Tuple<Guid, string, IOType, string>> sourceList, List<Tuple<Guid, string, IOType, string>> targetList)
        {
            return await Task.Run(() => Copy_varGUID(sourceList, targetList));
        }

        private List<Tuple<Guid, string, IOType, string>> Copy_varGUID(List<Tuple<Guid, string, IOType, string>> sourceList, List<Tuple<Guid, string, IOType, string>> targetList)
        {
            var returnList = new List<Tuple<Guid, string, IOType, string>>();

            foreach(var targetVar in targetList)
            {
                foreach(var sourceVar in sourceList)
                {
                    if (sourceVar.Item2 == targetVar.Item2)
                        returnList.Add(sourceVar);
                }
            }

            return returnList;
        }

        private async Task Create_slotVarConnectionAsync(string IOType, List<Guid> logicObjectGUIDs, List<Tuple<Guid, string, IOType, string>> variableList)
        {
            if(IOType == "Input")
            {
                foreach (var variable in variableList)
                {
                    await Api.Api.Interface.SendSlotVarAssignmentAsync(logicObjectGUIDs[0], "I_" + variable.Item2, variable.Item1);
                }    
            }
            else if(IOType == "InputStation")
            {
                foreach (var variable in variableList)
                {
                    var splitString = variable.Item2.Split("_");
                    await Api.Api.Interface.SendSlotVarAssignmentAsync(logicObjectGUIDs[0], "I_" + splitString[1], variable.Item1);
                }
            }
            else if (IOType == "Output")
            {
                foreach (var variable in variableList)
                {
                    await Api.Api.Interface.SendSlotVarAssignmentAsync(logicObjectGUIDs[0], "O_" + variable.Item2, variable.Item1);
                }
            }
            else if (IOType == "OutValue")
            {
                var counter = 0;
                foreach(var logicObjectGUID in logicObjectGUIDs)
                {
                    await Api.Api.Interface.SendSlotVarAssignmentAsync(logicObjectGUID, "OutValue", variableList[counter].Item1);

                    counter += 1;
                }
                    
            }
        }

        private async Task<string> Get_treeGUIDWithNameAsync(Guid objectGUID, string name)
        {
            return await Task.Run(() => Get_treeGUIDWithName(objectGUID, name));
        }

        private string Get_treeGUIDWithName(Guid objectGUID, string name)  // recursive method
        {
            var returnGUID = "";

            // check if the object name already contains the sought name
            if (Api.Api.Object.GetProperty(objectGUID, "Name").Contains(name))
            {
                returnGUID = objectGUID.ToString();
            }
            else
            {
                var childGUIDs = Api.Api.Object.GetAllChildrenFromSceneObject(objectGUID);
                foreach (var childGUID in childGUIDs)
                {
                    var property = Api.Api.Object.GetProperty(childGUID, "Name");
                    if (property.Contains(name))
                    {
                        returnGUID = childGUID;
                    }
                    else
                    {
                        if (Api.Api.Object.GetAllChildrenFromSceneObject(childGUID)[0] != "Object has no children")
                            returnGUID = Get_treeGUIDWithName(Guid.Parse(childGUID), name);
                    }
                    if (returnGUID != "")
                        break;
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

        public async Task ImportAsync(string filePath, string feeLibraryPath = null)
        {
            Status.Text = "import started";
            Status.Set_taskCount(31);

            // Find Fee library object GUIDs [Task 1]
            if (feeLibraryPath != null)
            {
                await Import_libraryGuidsAsync(feeLibraryPath);
            }
            else
            {
                Set_defaultLibraryGuids();
            }

            // Create Indexers
            var importDocument = CAEXDocument.LoadFromFile(filePath);
            var IH = importDocument.CAEXFile.InstanceHierarchy["DemonstratorProject"];
            var amlBRCL = AutomationMLBaseRoleClassLibType.RoleClassLib(importDocument);
            var RC = amlBRCL.ResourceStructure;
            var resourcesIE = AMLCustom.Get_IEs(IH, RC).First();

            var RCL = importDocument.CAEXFile.RoleClassLib["ResourceRoleClassLib"];
            RC = RCL.RoleClass["ProductionLine"];
            var plantIE = AMLCustom.Get_IEs(resourcesIE, RC).First();

            // Frame and TransportUnit [Task 2-3]
            RC = RCL.RoleClass["Frame"];
            var frameIE = AMLCustom.Get_IEs(plantIE, RC).First();
            var xPos = float.Parse(frameIE.Attribute["xPos"].Value,
                System.Globalization.CultureInfo.InvariantCulture) / 1000;
            var yPos = float.Parse(frameIE.Attribute["yPos"].Value,
                System.Globalization.CultureInfo.InvariantCulture) / 1000;
            var zPos = float.Parse(frameIE.Attribute["zPos"].Value,
                System.Globalization.CultureInfo.InvariantCulture) / 1000;

            string frameGUID = await Import_libraryObjectAsync("frame", name: "frame",
                xPos: xPos, yPos: yPos, zPos: zPos);  // import library object

            RC = RCL.RoleClass["TransportUnit"];
            var transportUnitIE = AMLCustom.Get_IEs(plantIE, RC).First();
            xPos = float.Parse(transportUnitIE.Attribute["xPos"].Value,
                System.Globalization.CultureInfo.InvariantCulture) / 1000;
            yPos = float.Parse(transportUnitIE.Attribute["yPos"].Value,
                System.Globalization.CultureInfo.InvariantCulture) / 1000;
            zPos = float.Parse(transportUnitIE.Attribute["zPos"].Value,
                System.Globalization.CultureInfo.InvariantCulture) / 1000;

            string transportUnitGUID = await Import_libraryObjectAsync("transportUnit", name: "transportUnit",
                xPos: xPos, yPos: yPos, zPos: zPos);  // import library object

            // Slots and Stations [Task 4-23]
            RC = RCL.RoleClass["Slot"];
            var slotIEs = AMLCustom.Get_IEs(plantIE, RC);

            var slotsAndStations = new List<string[]>();

            foreach (var slotIE in slotIEs)
            {
                var id = slotIE.Attribute["id"].Value;
                xPos = float.Parse(slotIE.Attribute["xPos"].Value,
                    System.Globalization.CultureInfo.InvariantCulture) / 1000;
                yPos = float.Parse(slotIE.Attribute["yPos"].Value,
                    System.Globalization.CultureInfo.InvariantCulture) / 1000;
                zPos = float.Parse(slotIE.Attribute["zPos"].Value,
                    System.Globalization.CultureInfo.InvariantCulture) / 1000;
                var side = slotIE.Attribute["side"].Value;
                var zRot = 0;
                if (side == "left")
                {
                    zRot = 180;
                }

                string slotGUID = await Import_libraryObjectAsync("slot", name: ("slot" + id),
                    xPos: xPos, yPos: yPos, zPos: zPos, zRot: zRot);  // import library object

                if (slotGUID != null)
                {
                    slotsAndStations.Add(new[] { id, slotGUID, "" });

                    if (slotIE.InternalElement.Exists)  // if slot has a station
                    {
                        var stationIE = slotIE.InternalElement[0];
                        var stationXPos = float.Parse(stationIE.Attribute["xPos"].Value,
                            System.Globalization.CultureInfo.InvariantCulture) / 1000;
                        var stationYPos = float.Parse(stationIE.Attribute["yPos"].Value,
                            System.Globalization.CultureInfo.InvariantCulture) / 1000;
                        var stationZPos = float.Parse(stationIE.Attribute["zPos"].Value,
                            System.Globalization.CultureInfo.InvariantCulture) / 1000;
                        var type = stationIE.Attribute["type"].Value;
                        var color = stationIE.Attribute["color"].Value;

                        // generate station name
                        var stationName = type;
                        stationName += "_" + color;

                        // import Library Object
                        string stationGUID = await Import_libraryObjectAsync(stationName, name: stationName,
                            xPos: xPos, yPos: yPos, zPos: zPos, zRot: zRot);  // import library object

                        if (stationGUID != null)
                            slotsAndStations.Last()[2] = stationGUID;
                    }
                    else
                        Status.Finish_task();
                }
            }

            // Arduinos [Task 24-25]
            string ssArduinoGUID = await Import_libraryObjectAsync("Logic_ss_arduino", name: "Logic_ss_arduino");  // import library object
            string arArduinoGUID = await Import_libraryObjectAsync("Logic_ar_arduino", name: "Logic_ar_arduino");  // import library object

            // Decompose(delete) the library root element [Task 26-31]
            
            await Remove_libraryRootElementsAsnyc();
            
            await Create_diagramAsync("kinematic_diagram", Guid.Parse(transportUnitGUID));
            await Create_diagramAsync("SS_arduino_diagram", Guid.Parse(ssArduinoGUID));
            await Create_diagramAsync("AR_arduino_diagram", Guid.Parse(arArduinoGUID));
            var lightBallGUID = Get_treeGUIDWithName(Guid.Parse(transportUnitGUID), "Lightball");
            await Create_diagramAsync("armLamp_diagram", Guid.Parse(lightBallGUID));
            var guidList = new List<Guid>();
            foreach(var slotAndStation in slotsAndStations)
            {
                if (slotAndStation[2] == "")
                    guidList.Add(Guid.Empty);
                else
                    guidList.Add(Guid.Parse(slotAndStation[2]));
            }
            await Create_diagramAsync("stationLamps_diagram", guidList);

            Status.Text = "import finished";
        }

#endregion

    }
}
