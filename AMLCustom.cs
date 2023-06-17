// Author: Yifan Wang

using Aml.Engine.CAEX;

namespace AMLConversion_ClassLibrary
{
    public static class AMLCustom
    {
        public static CAEXDocument Import_libraries(CAEXDocument targetDocument, CAEXDocument sourceDocument, List<string> IHImport = null, List<string> SUCLImport = null, List<string> RCLImport = null, List<string> ICLImport = null, List<string> ATLImport = null)
        {
            // Import InstanceHierarchy
            if (IHImport != null)
            {
                var sourceIHFiltered = sourceDocument.CAEXFile.InstanceHierarchy.Where(element => IHImport.Contains(element.Name));  // Pick the hierarchies equal to names in ImportList
                foreach (var sourceIH in sourceIHFiltered)
                {
                    targetDocument.CAEXFile.InstanceHierarchy.Insert(sourceIH);  // Import the picked instance hierarchies
                }
            }

            if (SUCLImport != null)
            {
                // Import SystemUnitClassLibrary
                var sourceSUCLFiltered = sourceDocument.CAEXFile.SystemUnitClassLib.Where(element => SUCLImport.Contains(element.Name));  // Pick the libraries equal to names in ImportList
                foreach (var sourceSUCL in sourceSUCLFiltered)
                {
                    targetDocument.CAEXFile.SystemUnitClassLib.Insert(sourceSUCL);  // Import the picked libraries
                }
            }

            if (RCLImport != null)
            {
                // Import RoleClassLibrary
                var sourceRCLFiltered = sourceDocument.CAEXFile.RoleClassLib.Where(element => RCLImport.Contains(element.Name));  // Pick the libraries equal to names in ImportList
                foreach (var sourceRCL in sourceRCLFiltered)
                {
                    targetDocument.CAEXFile.RoleClassLib.Insert(sourceRCL);  // Import the picked libraries
                }
            }

            if (ICLImport != null)
            {
                // Import InterfaceClassLibrary
                var sourceICLFiltered = sourceDocument.CAEXFile.InterfaceClassLib.Where(element => ICLImport.Contains(element.Name));  // Pick the libraries equal to names in ImportList
                foreach (var sourceICL in sourceICLFiltered)
                {
                    targetDocument.CAEXFile.InterfaceClassLib.Insert(sourceICL);  // Import the picked libraries
                }
            }

            // Import AttributeTypeLibrary
            if (ATLImport != null)
            {
                var sourceATLFiltered = sourceDocument.CAEXFile.AttributeTypeLib.Where(element => ATLImport.Contains(element.Name));  // Pick the libraries equal to names in ImportList
                foreach (var sourceATL in sourceATLFiltered)
                {
                    targetDocument.CAEXFile.AttributeTypeLib.Insert(sourceATL);    // Importy the picked libraries
                }
            }

            return targetDocument;
        }

        public static IEnumerable<SystemUnitFamilyType> Get_SUCs(SystemUnitClassLibType systemUnitClassLibrary, RoleFamilyType roleClass = null, string name = null)
        {
            IEnumerable<SystemUnitFamilyType> rSUCs = null;

            if (name != null)
            {
                rSUCs = systemUnitClassLibrary.Where(element => element.HasRoleClassReference(roleClass) == true && element.Name == name);
            }
            else if (name == null)
            {
                rSUCs = systemUnitClassLibrary.Where(element => element.HasRoleClassReference(roleClass) == true);
            }

            return rSUCs;
        }

        public static IEnumerable<InternalElementType> Get_IEs(IEnumerable<InternalElementType> internalElement, RoleFamilyType roleClass = null, SystemUnitFamilyType systemUnitClass = null, string attributeName = null, string attributeValue = null, string name = null)
        {
            IEnumerable<InternalElementType> rIEs = internalElement;

            if (systemUnitClass != null)
                rIEs = rIEs.Where(element => element.HasSystemUnitClassReference(systemUnitClass) == true);

            if (roleClass != null)
                rIEs = rIEs.Where(element => element.HasRoleClassReference(roleClass) == true);

            if (name != null)
                rIEs = rIEs.Where(element => element.Name == name);

            if (attributeName != null)
                rIEs = rIEs.Where(element => element.Attribute[attributeName].AttributeValue.ToString() == attributeValue);

            return rIEs;
        }

        public static IEnumerable<ExternalInterfaceType> Get_EIs(IEnumerable<ExternalInterfaceType> externalInterface, InterfaceFamilyType interfaceClass = null, string name = null)
        {
            IEnumerable<ExternalInterfaceType> rEIs = null;

            if (name != null)
            {
                if (interfaceClass == null)
                {
                    rEIs = externalInterface.Where(element => element.Name == name);
                }
                else if (interfaceClass != null)
                {
                    rEIs = externalInterface.Where(element => element.HasInterfaceClassReference(interfaceClass) == true && element.Name == name);
                }
            }
            else if (name == null)
            {
                rEIs = externalInterface.Where(element => element.HasInterfaceClassReference(interfaceClass) == true);
            }

            return rEIs;
        }

        public static InternalElementType Create_Instance(InternalElementType internalElement, SystemUnitFamilyType systemUnitClass = null, RoleFamilyType roleClass = null, bool asFirst = false, string instanceName = null)
        {
            InternalElementType rInternalElement = null;

            if (roleClass != null)
            {
                internalElement.Insert(roleClass.CreateClassInstance(), asFirst: asFirst);
                rInternalElement = internalElement;

            }
            else if (systemUnitClass != null)
            {
                internalElement.Insert(systemUnitClass.CreateClassInstance(), asFirst: asFirst);
                if (asFirst == true)
                {
                    rInternalElement = Get_IEs(internalElement, systemUnitClass: systemUnitClass).First();
                }
                else if (asFirst == false)
                {
                    rInternalElement = Get_IEs(internalElement, systemUnitClass: systemUnitClass).Last();
                }
                if (instanceName != null)
                {
                    rInternalElement.Name = instanceName;
                }
            }

            return rInternalElement;
        }
    }
}