// Author: Yifan Wang

using System.ComponentModel;

using Aml.Engine.CAEX;
using Aml.Engine.AmlObjects;

namespace AMLConversion_ClassLibrary
{
    public class Exporter_physicalPlant: ObservableObject
    {
        ///////////////////////////////////////////////////////////////////////////////////
        //
        // Constructor
        //
        ///////////////////////////////////////////////////////////////////////////////////
#region Constructor and attributes

        // Global variables
        private StatusInformation _status;
        private CAEXDocument _exportDocument;
        private AutomationMLBaseRoleClassLibType _amlBRCL;
        private AutomationMLInterfaceClassLibType _amlICL;
        private List<string[]> _slotProperties;
        private SerialCommunication _serial;
        
        public string productSinkTypeName = "productSink";
        public string productSourceTypeName = "productSource";

        public Exporter_physicalPlant(SerialCommunication ser = null)  // Constructor
        {
            _status = new StatusInformation();
            _status.PropertyChanged += StatusOnPropertyChanged;
            _serial = ser;
            _slotProperties = new List<string[]>();
    }

        public SerialCommunication Serial
        {
            get => _serial;
            set
            {
                _serial = value;
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

        public List<string[]> SlotProperties
        {
            get => _slotProperties;
            set
            {
                _slotProperties = value;
                OnPropertyChanged("SlotProperties");
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
#region preperation functions

        private CAEXDocument Prepare(CAEXDocument targetDocument, CAEXDocument sourceDocument, bool ih = true, bool sucl = true, bool rcl = true, bool icl = true, bool atl = true)
        {
            Status.Text = "prepare AML document";

            // Generate metatdata
            SourceDocumentInformationType documentInformation = targetDocument.CAEXFile.SourceDocumentInformation.First();
            documentInformation.OriginID = Guid.NewGuid().ToString();
            documentInformation.OriginName = "physischerDemonstrator";
            documentInformation.OriginVersion = "0.1";
            documentInformation.OriginVendor = "TH Köln";
            documentInformation.OriginVendorURL = "www.th-koeln.de";
            documentInformation.OriginProjectID = "Demonstrator";
            documentInformation.OriginProjectTitle = "Demonstrator";
            documentInformation.LastWritingDateTime = DateTime.Now;

            // Import blueprintLibraries
            var IHImport = new List<string>();
            var SUCLImport = new List<string>();
            var RCLImport = new List<string>();
            var ICLImport = new List<string>();
            var ATLImport = new List<string>();

            if (ih)
                IHImport.AddRange(new List<string> { "DemonstratorProject" });
            if (sucl)
                SUCLImport.AddRange(new List<string> { "Resourcemodel", "Processmodel", "Productmodel" });
            if (rcl)
                RCLImport.AddRange(new List<string> { "ResourceRoleClassLib", "ProcessRoleClassLib", "ProductRoleClassLib" });
            if (icl)
                ICLImport.AddRange(new List<string> { "PPRConnectorInterfaceClassLib" });
            if (atl)
                ATLImport.AddRange(new List<string> { "AttributeTypeLib" });
            targetDocument = AMLCustom.Import_libraries(targetDocument, sourceDocument, IHImport: IHImport, SUCLImport: SUCLImport, RCLImport: RCLImport, ICLImport: ICLImport, ATLImport: ATLImport);

            // adds the base libraries to the document
            _amlICL = AutomationMLInterfaceClassLibType.InterfaceClassLib(targetDocument);
            _amlBRCL = AutomationMLBaseRoleClassLibType.RoleClassLib(targetDocument);

            Status.Finish_task();

            return targetDocument;
        }

        private List<string[]> Import_slotPropertiesFromCSV(string CSVPath)
        {
            Status.Text = "import slotProperties from CSV";

            // resets the slotProperties list
            var list = new List<string[]>();

            // import csv data
            using (var csvFile = new StreamReader(CSVPath))
            {
                var i = 0;
                while (!csvFile.EndOfStream)
                {
                    var line = csvFile.ReadLine();
                    var values = line.Split(';');

                    if (i > 0)  // Skip header
                    {
                        list.Add(values);
                    }
                    i += 1;
                }
            }

            Status.Finish_task();

            return list;
        }

        private List<string[]> Refresh_slotProperties(CAEXDocument document, List<string[]> slotProperties)
        {
            Status.Text = "refresh slotProperties";

            // Create indexers
            var RCLResources = document.CAEXFile.RoleClassLib["ResourceRoleClassLib"];
            var IH = document.CAEXFile.InstanceHierarchy["DemonstratorProject"];
            var RC = _amlBRCL.ResourceStructure;
            var resourceIE = AMLCustom.Get_IEs(IH, roleClass: RC).First();
            RC = RCLResources.RoleClass["ProductionLine"];
            var plantIE = AMLCustom.Get_IEs(resourceIE, roleClass: RC).First();

            // Search for slots
            RC = RCLResources.RoleClass["Slot"];
            var slotIEs = plantIE.InternalElement.Where(element => element.HasRoleClassReference(RC));

            // Resets the slotProperties List
            slotProperties = new List<string[]>();

            foreach (var slotIE in slotIEs)
            {
                var id = slotIE.Attribute["id"].AttributeValue.ToString();  // id value as string
                var slotProperty = new []{ id, "", "" };

                if (slotIE.InternalElement.Count > 0)  // if slot has a station
                {
                    // Indexers
                    var stationIE = slotIE.InternalElement[0];
                    var color = stationIE.Attribute["color"].AttributeValue.ToString();  // color value as string
                    var type = stationIE.Attribute["type"].AttributeValue.ToString();  // type value as string
                    
                    // Replace the array values
                    slotProperty[1] = color;
                    slotProperty[2] = type;
                }
                // Add new slotProperties list entry
                slotProperties.Add(slotProperty);
            }

            Status.Finish_task();

            return slotProperties;
        }

#endregion

#region resources

        private CAEXDocument Generate_resources(CAEXDocument document)
        {
            Status.Text = "generate resources";

            // Create indexers
            var SUCResources = document.CAEXFile.SystemUnitClassLib["Resourcemodel"];
            var RCLResources = document.CAEXFile.RoleClassLib["ResourceRoleClassLib"];

            // Create Instance Hierarchy
            var IH = document.CAEXFile.InstanceHierarchy.Append("DemonstratorProject");

            // Create ResourceStructure InternalElement
            var resourceIE = IH.InternalElement.Append("Resources");  // Create InternalElement
            var BRC = _amlBRCL.ResourceStructure;  // Indexer for BaseRoleClass
            AMLCustom.Create_Instance(resourceIE, roleClass: BRC, asFirst: false);  // RoleClass assignment

            // Create ProductionLine InternalElement
            var plantIE = resourceIE.InternalElement.Append("Demonstrator");  // Create InternalElement
            var RC = RCLResources.RoleClass["ProductionLine"];  // Indexer for RoleClass
            AMLCustom.Create_Instance(plantIE, roleClass: RC, asFirst: false);  // RoleClass assignment

            // Create Frame InternalElement from SystemUnitClass
            var SUC = SUCResources.SystemUnitClass["Frame"];  // Indexer for SystemUnitClass
            AMLCustom.Create_Instance(plantIE, systemUnitClass: SUC, asFirst: false);   // Instance creation

            // Create TransportUnit InternalElement from SystemUnitClass
            SUC = SUCResources.SystemUnitClass["TransportUnit"];
            AMLCustom.Create_Instance(plantIE, systemUnitClass: SUC, asFirst: false);

            // Create numerous Slot (and Station) InternalElements from SystemUnitClass
            SUC = SUCResources.SystemUnitClass["Slot"];
            var SUC2 = SUCResources.SystemUnitClass["Station"];

            var j = 0;
            foreach (var line in _slotProperties)  // for every Slot
            {
                var slotIE = AMLCustom.Create_Instance(plantIE, systemUnitClass: SUC, asFirst: false);  // Create Slot instance

                // Slot: id attribute
                var attr = slotIE.Attribute["id"];  // indexer for AttributeValue "id"
                attr.AttributeValue = j + 1;  // assign attribute value with automatic typeconversion to assigned AttributeDataType

                // Slot: side and position attributes
                var sideAttr = slotIE.Attribute["side"];
                var xPosAttr = slotIE.Attribute["xPos"];
                var yPosAttr = slotIE.Attribute["yPos"];
                var zPosAttr = slotIE.Attribute["zPos"];
                var factor = 0;

                if ((j + 1) % 2 == 0)  // if "id" is straight
                {
                    sideAttr.AttributeValue = "left";
                    yPosAttr.AttributeValue = 413.2;
                    factor = (j - 1);
                }
                else  // if "id" is odd
                {
                    sideAttr.AttributeValue = "right";
                    yPosAttr.AttributeValue = 186.8;
                    factor = j;
                }
                xPosAttr.AttributeValue = 150 + factor * 62.5;
                zPosAttr.AttributeValue = 99;

                if (line[2] != "empty" && line[2] != "")  // check if slot has a station
                {
                    var stationIE = AMLCustom.Create_Instance(slotIE, systemUnitClass: SUC2, asFirst: false); // Create Station instance

                    // Station: color attribute
                    attr = stationIE.Attribute["color"];
                    attr.AttributeValue = line[1];  // assign attribute value

                    // Station: type attribute
                    attr = stationIE.Attribute["type"];
                    attr.AttributeValue = line[2];  // assign attribute value

                    // Station: position attributes
                    var xPosStationAttr = stationIE.Attribute["xPos"];
                    var yPosStationAttr = stationIE.Attribute["yPos"];
                    var zPosStationAttr = stationIE.Attribute["zPos"];

                    xPosStationAttr.AttributeValue = slotIE.Attribute["xPos"].AttributeValue;  // same as slot attribute
                    yPosStationAttr.AttributeValue = slotIE.Attribute["yPos"].AttributeValue;  // same as slot attribute
                    zPosStationAttr.AttributeValue = 109;
                }

                j++;
            }

            Status.Finish_task();

            return document;
        }

#endregion

#region processes

        private List<InternalElementType[]> Generate_transportList(List<InternalElementType> productSinkSlots, List<InternalElementType> productSourceSlots, string rule, List<string> colorList)
        {
            var transportList = new List<InternalElementType[]>();

            if (rule == "automatic")
            {
                foreach (var productSinkSlot in productSinkSlots.ToList())  // ToList to update the list with LINQ query
                {
                    var productSinkStation = productSinkSlot.InternalElement[0];
                    foreach (var productSourceSlot in productSourceSlots.ToList())  // ToList to update the list with LINQ query
                    {
                        var productSourceStation = productSourceSlot.InternalElement[0];
                        var productSinkColor = productSinkStation.Attribute["color"].AttributeValue.ToString();
                        var productSourceColor = productSourceStation.Attribute["color"].AttributeValue.ToString();
                        if (productSinkColor == productSourceColor)
                        {
                            var transportPair = new InternalElementType[2] { productSourceSlot, productSinkSlot };
                            transportList.Add(transportPair);
                            productSourceSlots.Remove(productSourceSlot);
                            break;
                        }
                    }
                    productSinkSlots.Remove(productSinkSlot);
                }
            }
            else if (rule == "custom")
            {
                foreach (var color in colorList)
                {
                    var colorFound = false;

                    foreach (var productSinkSlot in productSinkSlots.ToList())  // ToList to update the list with LINQ query
                    {
                        var productSinkStation = productSinkSlot.InternalElement[0];
                        var productSinkColor = productSinkStation.Attribute["color"].AttributeValue.ToString();
                        if (color == productSinkColor)
                        {
                            foreach (var productSourceSlot in productSourceSlots.ToList())  // ToList to update the list with LINQ query
                            {
                                var productSourceStation = productSourceSlot.InternalElement[0];
                                var productSourceColor = productSourceStation.Attribute["color"].AttributeValue.ToString();
                                if (color == productSourceColor)
                                {
                                    var transportPair = new InternalElementType[2] { productSourceSlot, productSinkSlot };
                                    transportList.Add(transportPair);
                                    productSourceSlots.Remove(productSourceSlot);
                                    colorFound = true;
                                }
                                if(colorFound)
                                    break;
                            }
                            productSinkSlots.Remove(productSinkSlot);
                        }
                        if (colorFound)
                            break;
                    }
                }
            }

            return transportList;
        }

        private CAEXDocument Generate_processes(CAEXDocument document, string rule, List<string> colorList = null)
        {
            Status.Text = "generate processes";

            // Create Indexers
            var IH = document.CAEXFile.InstanceHierarchy[0];
            var resourceStructureRC = _amlBRCL.ResourceStructure;
            var resourceIE = AMLCustom.Get_IEs(IH, roleClass: resourceStructureRC).First();
            var resourceRCL = document.CAEXFile.RoleClassLib["ResourceRoleClassLib"];
            var plantRC = resourceRCL.RoleClass["ProductionLine"];
            var plantIE = AMLCustom.Get_IEs(resourceIE, roleClass: plantRC).First();
            var slotRC = resourceRCL.RoleClass["Slot"];
            var slotIEs = AMLCustom.Get_IEs(plantIE, roleClass: slotRC);
            var productStructureRC = _amlBRCL.ProductStructure;
            var productIE = AMLCustom.Get_IEs(IH, roleClass: productStructureRC).First();
            var productRCL = document.CAEXFile.RoleClassLib["ProductRoleClassLib"];

            // Pick the Slots with stations
            var slotsWithStationFiltered = slotIEs.Where(element => element.InternalElement.Count > 0);

            // Pick the Slots with Stations with "productSink" type
            var productSinkSlotsFiltered = slotsWithStationFiltered.Where(element => element.InternalElement[0].Attribute["type"].AttributeValue.ToString() == productSinkTypeName).ToList();
            foreach (var productSinkSlot in productSinkSlotsFiltered)
            {
                var productSink = productSinkSlot.InternalElement[0];
            }

            // Pick the Slots with Stations with "productSource" type
            var productSourceSlotsFiltered = slotsWithStationFiltered.Where(element => element.InternalElement[0].Attribute["type"].AttributeValue.ToString() == productSourceTypeName).ToList();
            foreach (var productSourceSlot in productSourceSlotsFiltered)
            {
                var productSource = productSourceSlot.InternalElement[0];
            }

            // Create Indexers
            var processRCL = document.CAEXFile.RoleClassLib["ProcessRoleClassLib"];
            var processStructureRC = _amlBRCL.ProcessStructure;
            var SUCLProcesses = document.CAEXFile.SystemUnitClassLib["Processmodel"];
            var RCLProcesses = document.CAEXFile.RoleClassLib["ProcessRoleClassLib"];

            var transportUnitRC = resourceRCL.RoleClass["TransportUnit"];
            var transportUnitIE = AMLCustom.Get_IEs(plantIE, roleClass: transportUnitRC).First();
            var sledgeRC = resourceRCL.RoleClass["Sledge"];
            var sledgeIE = AMLCustom.Get_IEs(transportUnitIE, roleClass: sledgeRC).First();
            var armRC = resourceRCL.RoleClass["Arm"];
            var armIE = AMLCustom.Get_IEs(transportUnitIE, armRC).First();
            var pprICL = document.CAEXFile.InterfaceClassLib["PPRConnectorInterfaceClassLib"];
            var parentRC = processRCL.RoleClass["ProductionStep"];

            // Create ProcessStructure InternalElement
            var processIEDummy = IH.InternalElement.Append("Processes");  // Create InternalElement
            var processIE = IH.InternalElement.InsertAt(1, (InternalElementType)processIEDummy.Copy());  // Copy InternalElement to 2nd position (index 1)
            IH.InternalElement.RemoveElement(processIEDummy);  // delete dummy Internal Element
            var BRC = _amlBRCL.ProcessStructure;  // Indexer for BaseRoleClass
            processIE.Insert(BRC.CreateClassInstance());  // RoleClass assignment

            // Create Job Init from SystemUnitClass
            var RC = processRCL.RoleClass["Job"];
            var SUC = AMLCustom.Get_SUCs(SUCLProcesses, roleClass: RC, "Init").First();  // Indexer for SystemUnitClass
            var jobIE = AMLCustom.Create_Instance(processIE, SUC);   // Instance creation
            RC = processRCL.RoleClass["ProductionProcess"];
            var productionProcessIE = AMLCustom.Get_IEs(jobIE, roleClass: RC).First();
            RC = parentRC.RoleClass["MoveToSide"];
            var productionStepIE = AMLCustom.Get_IEs(productionProcessIE, roleClass: RC).First();
            var rpIC = pprICL.InterfaceClass["Connector_ResourceProcess"];
            var eppcIC = pprICL.InterfaceClass["Connector_EndproductProcess"];
            var wpIC = pprICL.InterfaceClass["Connector_WorkpieceProcess"];

            var processEI = AMLCustom.Get_EIs(productionStepIE.ExternalInterface, interfaceClass: rpIC).First();
            var resourceEI = AMLCustom.Get_EIs(armIE.ExternalInterface, interfaceClass: rpIC).First();
            var relation = InternalLinkType.New_InternalLink(processEI, resourceEI, "pprRelation");  // create pprRelation Resource-Process

            RC = parentRC.RoleClass["Reference"];
            productionStepIE = AMLCustom.Get_IEs(productionProcessIE, roleClass: RC).First();

            processEI = AMLCustom.Get_EIs(productionStepIE.ExternalInterface, interfaceClass: rpIC).First();
            resourceEI = AMLCustom.Get_EIs(sledgeIE.ExternalInterface, interfaceClass: rpIC).First();
            relation = InternalLinkType.New_InternalLink(processEI, resourceEI, "pprRelation");  // create pprRelation Resource-Process

            RC = parentRC.RoleClass["Home"];
            productionStepIE = AMLCustom.Get_IEs(productionProcessIE, roleClass: RC).First();

            processEI = AMLCustom.Get_EIs(productionStepIE.ExternalInterface, interfaceClass: rpIC).First();
            resourceEI = AMLCustom.Get_EIs(sledgeIE.ExternalInterface, interfaceClass: rpIC).First();
            relation = InternalLinkType.New_InternalLink(processEI, resourceEI, "pprRelation");  // create pprRelation Resource-Process

            RC = parentRC.RoleClass["LEDCheck"];
            productionStepIE = AMLCustom.Get_IEs(productionProcessIE, roleClass: RC).First();

            processEI = AMLCustom.Get_EIs(productionStepIE.ExternalInterface, interfaceClass: rpIC).First();
            resourceEI = AMLCustom.Get_EIs(armIE.ExternalInterface, interfaceClass: rpIC).First();
            relation = InternalLinkType.New_InternalLink(processEI, resourceEI, "pprRelation");  // create pprRelation Resource-Process

            // Create Job Order_Customer1 from SystemUnitClass
            RC = processRCL.RoleClass["Job"];
            SUC = AMLCustom.Get_SUCs(SUCLProcesses, roleClass: RC, "Jobname").First();
            var orderIE = AMLCustom.Create_Instance(processIE, SUC);
            orderIE.Name = "Order_Customer1";

            // Generate transportList
            var transportList = Generate_transportList(productSinkSlotsFiltered, productSourceSlotsFiltered, rule, colorList);
            
            // Process transportList
            foreach (var pair in transportList)
            {
                // Create Indexers
                var departureSlotIE = pair[0];
                var destinationSlotIE = pair[1];
                var departureStationIE = departureSlotIE.InternalElement[0];
                var destinationStationIE = destinationSlotIE.InternalElement[0];
                var departureID = departureSlotIE.Attribute["id"].AttributeValue;
                var destinationID = destinationSlotIE.Attribute["id"].AttributeValue;
                var departureColor = departureStationIE.Attribute["color"].AttributeValue;
                var destinationColor = destinationStationIE.Attribute["color"].AttributeValue;

                // Create InternalElement Transport
                RC = processRCL.RoleClass["ProductionProcess"];
                SUC = AMLCustom.Get_SUCs(SUCLProcesses, roleClass: RC, "Transport").First();
                var transportIE = AMLCustom.Create_Instance(orderIE, systemUnitClass: SUC);
                transportIE.Attribute["DepartureID"].AttributeValue = departureID;
                transportIE.Attribute["DestinationID"].AttributeValue = destinationID;
                processEI = AMLCustom.Get_EIs(transportIE.ExternalInterface, interfaceClass: eppcIC).First();
                RC = productRCL.RoleClass["Endproduct"];
                var endproductIE = AMLCustom.Get_IEs(productIE, roleClass: RC, attributeName: "color", attributeValue: departureColor.ToString()).First();
                var productEI = AMLCustom.Get_EIs(endproductIE.ExternalInterface, interfaceClass: eppcIC).First();
                relation = InternalLinkType.New_InternalLink(processEI, productEI, "pprRelation");  // create pprRelation Endproduct-Process

                RC = parentRC.RoleClass["MoveTo"];
                productionStepIE = AMLCustom.Get_IEs(transportIE, roleClass: RC).First();
                productionStepIE.Attribute["slotID"].AttributeValue = departureID;
                processEI = AMLCustom.Get_EIs(productionStepIE.ExternalInterface, interfaceClass: rpIC).First();
                resourceEI = AMLCustom.Get_EIs(sledgeIE.ExternalInterface, interfaceClass: rpIC).First();
                relation = InternalLinkType.New_InternalLink(processEI, resourceEI, "pprRelation");  // create pprRelation Resource-Process

                productionStepIE = AMLCustom.Get_IEs(transportIE, roleClass: RC).ElementAt(1);
                productionStepIE.Attribute["slotID"].AttributeValue = destinationID;
                processEI = AMLCustom.Get_EIs(productionStepIE.ExternalInterface, interfaceClass: rpIC).First();
                resourceEI = AMLCustom.Get_EIs(sledgeIE.ExternalInterface, interfaceClass: rpIC).First();
                relation = InternalLinkType.New_InternalLink(processEI, resourceEI, "pprRelation");  // create pprRelation Resource-Process

                RC = parentRC.RoleClass["MoveToSide"];
                productionStepIE = AMLCustom.Get_IEs(transportIE, roleClass: RC).First();
                productionStepIE.Attribute["slotID"].AttributeValue = departureID;
                processEI = AMLCustom.Get_EIs(productionStepIE.ExternalInterface, interfaceClass: rpIC).First();
                resourceEI = AMLCustom.Get_EIs(armIE.ExternalInterface, interfaceClass: rpIC).First();
                relation = InternalLinkType.New_InternalLink(processEI, resourceEI, "pprRelation");  // create pprRelation Resource-Process

                productionStepIE = AMLCustom.Get_IEs(transportIE, roleClass: RC).ElementAt(1);
                processEI = AMLCustom.Get_EIs(productionStepIE.ExternalInterface, interfaceClass: rpIC).First();
                resourceEI = AMLCustom.Get_EIs(armIE.ExternalInterface, interfaceClass: rpIC).First();
                relation = InternalLinkType.New_InternalLink(processEI, resourceEI, "pprRelation");  // create pprRelation Resource-Process

                productionStepIE = AMLCustom.Get_IEs(transportIE, roleClass: RC).ElementAt(2);
                productionStepIE.Attribute["slotID"].AttributeValue = destinationID;
                processEI = AMLCustom.Get_EIs(productionStepIE.ExternalInterface, interfaceClass: rpIC).First();
                resourceEI = AMLCustom.Get_EIs(armIE.ExternalInterface, interfaceClass: rpIC).First();
                relation = InternalLinkType.New_InternalLink(processEI, resourceEI, "pprRelation");  // create pprRelation Resource-Process

                productionStepIE = AMLCustom.Get_IEs(transportIE, roleClass: RC).ElementAt(3);
                processEI = AMLCustom.Get_EIs(productionStepIE.ExternalInterface, interfaceClass: rpIC).First();
                resourceEI = AMLCustom.Get_EIs(armIE.ExternalInterface, interfaceClass: rpIC).First();
                relation = InternalLinkType.New_InternalLink(processEI, resourceEI, "pprRelation");  // create pprRelation Resource-Process

                RC = parentRC.RoleClass["TakeProductAsGripper"];
                productionStepIE = AMLCustom.Get_IEs(transportIE, roleClass: RC).First();
                productionStepIE.Attribute["color"].AttributeValue = departureColor;
                processEI = AMLCustom.Get_EIs(productionStepIE.ExternalInterface, interfaceClass: rpIC).First();
                resourceEI = AMLCustom.Get_EIs(armIE.ExternalInterface, interfaceClass: rpIC).First();
                relation = InternalLinkType.New_InternalLink(processEI, resourceEI, "pprRelation");  // create pprRelation Resource-Process
                processEI = AMLCustom.Get_EIs(productionStepIE.ExternalInterface, interfaceClass: wpIC).First();
                RC = productRCL.RoleClass["Workpiece"];
                var workpieceIE = AMLCustom.Get_IEs(productIE, roleClass: RC, attributeName: "color", attributeValue: departureColor.ToString()).First();
                productEI = AMLCustom.Get_EIs(workpieceIE.ExternalInterface, interfaceClass: wpIC).First();
                relation = InternalLinkType.New_InternalLink(processEI, productEI, "pprRelation");  // create pprRelation Workpiece-Process

                RC = parentRC.RoleClass["DropProductAsGripper"];
                productionStepIE = AMLCustom.Get_IEs(transportIE, roleClass: RC).First();
                processEI = AMLCustom.Get_EIs(productionStepIE.ExternalInterface, interfaceClass: rpIC).First();
                resourceEI = AMLCustom.Get_EIs(armIE.ExternalInterface, interfaceClass: rpIC).First();
                relation = InternalLinkType.New_InternalLink(processEI, resourceEI, "pprRelation");  // create pprRelation Resource-Process
                processEI = AMLCustom.Get_EIs(productionStepIE.ExternalInterface, interfaceClass: wpIC).First();
                productEI = AMLCustom.Get_EIs(workpieceIE.ExternalInterface, interfaceClass: wpIC).First();
                relation = InternalLinkType.New_InternalLink(processEI, productEI, "pprRelation");  // create pprRelation Workpiece-Process

                RC = parentRC.RoleClass["TakeProductAsStation"];
                productionStepIE = AMLCustom.Get_IEs(transportIE, roleClass: RC).First();
                productionStepIE.Attribute["slotID"].AttributeValue = destinationID;
                productionStepIE.Attribute["color"].AttributeValue = departureColor;
                processEI = AMLCustom.Get_EIs(productionStepIE.ExternalInterface, interfaceClass: rpIC).First();
                resourceEI = AMLCustom.Get_EIs(destinationStationIE.ExternalInterface, interfaceClass: rpIC).First();
                relation = InternalLinkType.New_InternalLink(processEI, resourceEI, "pprRelation");  // create pprRelation Resource-Process
                processEI = AMLCustom.Get_EIs(productionStepIE.ExternalInterface, interfaceClass: wpIC).First();
                productEI = AMLCustom.Get_EIs(workpieceIE.ExternalInterface, interfaceClass: wpIC).First();
                relation = InternalLinkType.New_InternalLink(processEI, productEI, "pprRelation");  // create pprRelation Workpiece-Process

                RC = parentRC.RoleClass["DropProductAsStation"];
                productionStepIE = AMLCustom.Get_IEs(transportIE, roleClass: RC).First();
                productionStepIE.Attribute["slotID"].AttributeValue = departureID;
                processEI = AMLCustom.Get_EIs(productionStepIE.ExternalInterface, interfaceClass: rpIC).First();
                resourceEI = AMLCustom.Get_EIs(destinationStationIE.ExternalInterface, interfaceClass: rpIC).First();
                relation = InternalLinkType.New_InternalLink(processEI, resourceEI, "pprRelation");  // create pprRelation Resource-Process
                processEI = AMLCustom.Get_EIs(productionStepIE.ExternalInterface, interfaceClass: wpIC).First();
                productEI = AMLCustom.Get_EIs(workpieceIE.ExternalInterface, interfaceClass: wpIC).First();
                relation = InternalLinkType.New_InternalLink(processEI, productEI, "pprRelation");  // create pprRelation Workpiece-Process
            }

            // Create Internal Element Park
            RC = processRCL.RoleClass["ProductionProcess"];
            SUC = AMLCustom.Get_SUCs(SUCLProcesses, roleClass: RC, "Park").First();
            var parkIE = AMLCustom.Create_Instance(orderIE, systemUnitClass: SUC);

            RC = parentRC.RoleClass["DropProductAsGripper"];
            productionStepIE = AMLCustom.Get_IEs(parkIE, roleClass: RC).First();
            processEI = AMLCustom.Get_EIs(productionStepIE.ExternalInterface, interfaceClass: rpIC).First();
            resourceEI = AMLCustom.Get_EIs(armIE.ExternalInterface, interfaceClass: rpIC).First();
            relation = InternalLinkType.New_InternalLink(processEI, resourceEI, "pprRelation");  // create pprRelation Resource-Process

            RC = parentRC.RoleClass["MoveToSide"];
            productionStepIE = AMLCustom.Get_IEs(parkIE, roleClass: RC).First();
            processEI = AMLCustom.Get_EIs(productionStepIE.ExternalInterface, interfaceClass: rpIC).First();
            resourceEI = AMLCustom.Get_EIs(armIE.ExternalInterface, interfaceClass: rpIC).First();
            relation = InternalLinkType.New_InternalLink(processEI, resourceEI, "pprRelation");  // create pprRelation Resource-Process

            RC = parentRC.RoleClass["Park"];
            productionStepIE = AMLCustom.Get_IEs(parkIE, roleClass: RC).First();
            processEI = AMLCustom.Get_EIs(productionStepIE.ExternalInterface, interfaceClass: rpIC).First();
            resourceEI = AMLCustom.Get_EIs(sledgeIE.ExternalInterface, interfaceClass: rpIC).First();
            relation = InternalLinkType.New_InternalLink(processEI, resourceEI, "pprRelation");  // create pprRelation Resource-Process

            RC = parentRC.RoleClass["DropProductAsStation"];
            var productionStepIEs = AMLCustom.Get_IEs(parkIE, roleClass: RC);
            var idCounter = 1;
            foreach (var IE in productionStepIEs)
            {
                processEI = AMLCustom.Get_EIs(IE.ExternalInterface, interfaceClass: rpIC).First();
                var stationIE = slotIEs.ElementAt(idCounter - 1).InternalElement[0];
                if (stationIE != null)  // if slot has a station
                {
                    resourceEI = AMLCustom.Get_EIs(stationIE.ExternalInterface, interfaceClass: rpIC).First();
                    relation = InternalLinkType.New_InternalLink(processEI, resourceEI, "pprRelation");  // create pprRelation Resource-Process
                }

                idCounter += 1;
            }

            Status.Finish_task();

            return document;
        }

#endregion

#region products

        private CAEXDocument Generate_products(CAEXDocument document)
        {
            Status.Text = "generate products";

            // Create Indexers
            var IH = document.CAEXFile.InstanceHierarchy[0];
            var productStructureRC = _amlBRCL.ProductStructure;
            var productRCL = document.CAEXFile.RoleClassLib["ProductRoleClassLib"];
            var productSUCL = document.CAEXFile.SystemUnitClassLib["Productmodel"];
            List<string> productTypes = new List<string>();

            // Create ProductStructure InternalElement
            var productIE = IH.InternalElement.Append("Products");  // Create InternalElement
            AMLCustom.Create_Instance(productIE, roleClass: productStructureRC, asFirst: false);  // RoleClass assignment

            // detect all productTypes
            foreach (var slot in _slotProperties)
            {
                var productType = slot[1];
                if (!productTypes.Contains(productType) && productType != "")
                {
                    productTypes.Add(productType);
                }
            }

            // generate productTypes and related workpieceTypes
            foreach (var pType in productTypes)
            {
                var RC = productRCL.RoleClass["Workpiece"];
                var SUC = AMLCustom.Get_SUCs(productSUCL, roleClass: RC).First();
                var workpieceIE = AMLCustom.Create_Instance(productIE, systemUnitClass: SUC);  // Instance creation
                var workpieceAttribute = workpieceIE.Attribute["color"];
                workpieceAttribute.AttributeValue = pType;  // assign attribute value

                RC = productRCL.RoleClass["Endproduct"];
                SUC = AMLCustom.Get_SUCs(productSUCL, roleClass: RC).First();
                var endproductIE = AMLCustom.Create_Instance(productIE, systemUnitClass: SUC);  // Instance creation
                var endproductAttr = endproductIE.Attribute["color"];
                endproductAttr.AttributeValue = pType;  // assign attribute value

                switch (pType)
                {
                    case "red":
                        workpieceIE.Name = "RedUnprocessed";
                        endproductIE.Name = "RedProduct";
                        break;
                    case "blue":
                        workpieceIE.Name = "BlueUnprocessed";
                        endproductIE.Name = "BlueProduct";
                        break;
                    case "green":
                        workpieceIE.Name = "GreenUnprocessed";
                        endproductIE.Name = "GreenProduct";
                        break;
                }
            }

            Status.Finish_task();

            return document;
        }
        #endregion

        ///////////////////////////////////////////////////////////////////////////////////
        //
        // Public functions
        //
        ///////////////////////////////////////////////////////////////////////////////////
#region Public functions

        public void Save_file(string targetPath)
        {
            Status.Text = "save file";
            Status.Set_taskCount(1);

            _exportDocument.SaveToFile(targetPath, true);  // Save caex document as file

            Status.Finish_task();
            Status.Text = "file saved";
        }

        public async Task Export_topologyAsync(string blueprintPath, string CSVPath = null)
        {
            await Task.Run(() => Export_topology(blueprintPath, CSVPath));
        }

        public void Export_topology(string blueprintPath, string CSVPath = null)
        {
            // Create report event
            Status.Text = "create AML document";
            Status.Progress = 0;

            // Create indexers
            var topologyDocument = CAEXDocument.New_CAEXDocument();
            var blueprintDocument = CAEXDocument.LoadFromFile(blueprintPath);

            Status.Set_taskCount(3);
            topologyDocument = Prepare(topologyDocument, blueprintDocument, ih: false);
            
            if (CSVPath != null)
            {
                _slotProperties = Import_slotPropertiesFromCSV(CSVPath);
            }

            topologyDocument = Generate_resources(topologyDocument);
            _exportDocument = topologyDocument;

            // Create report event
            Status.Text = "AML document created";
            
        }

        public async Task Export_plantfileAsync(string topologyPath, string rule, List<string> colorList = null)
        {
            await Task.Run(() => Export_plantfile(topologyPath, rule, colorList));
        }

        public void Export_plantfile(string topologyPath, string rule, List<string> colorList = null)
        {
            // Create report event
            Status.Text = "Create AML document";
            Status.Progress = 0;

            Status.Set_taskCount(4);
            var plantmodelDocument = CAEXDocument.New_CAEXDocument();
            var topologyDocument = CAEXDocument.LoadFromFile(topologyPath);
            plantmodelDocument = Prepare(plantmodelDocument, topologyDocument);
            _slotProperties = Refresh_slotProperties(plantmodelDocument, _slotProperties);
            plantmodelDocument = Generate_products(plantmodelDocument);
            plantmodelDocument = Generate_processes(plantmodelDocument, rule, colorList: colorList);
            
            _exportDocument = plantmodelDocument;

            // Create report event
            Status.Text = "AML document created";
        }

#endregion

    }
}
