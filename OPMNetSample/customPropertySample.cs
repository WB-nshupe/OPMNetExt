using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows.OPM;
using System;
using System.Reflection;
using System.Runtime.InteropServices;
using Autodesk.AutoCAD.DatabaseServices;

namespace OPMNetSample
{
    #region Our Custom Property

    #region COM Stuff

    [
        Guid("F60AE3DA-0373-4d24-82D2-B2646517ABCB"),
        ProgId("OPMNetSample.CustomProperty.1"),

        // No class interface is generated for this class and
        // no interface is marked as the default.
        // Users are expected to expose functionality through
        // interfaces that will be explicitly exposed by the object
        // This means the object can only expose interfaces we define

        ClassInterface(ClassInterfaceType.None),
        // Set the default COM interface that will be used for
        // Automation. Languages like: C#, C++ and VB allow to 
        //query for interface's we're interested in but Automation 
        // only aware languages like javascript do not allow to 
        // query interface(s) and create only the default one

        ComDefaultInterface(typeof(IDynamicProperty2)),
        ComVisible(true)
    ]

    #endregion
    public class CustomProp : IDynamicProperty2, ICategorizeProperties {
        private IDynamicPropertyNotify2 m_pSink =null;

       // Unique property ID
        void IDynamicProperty2.GetGUID(out Guid propGUID)
        {
            propGUID = new Guid("F60AE3DA-0373-4d24-82D2-B2646517ABCB");
        }

        // Property display nameD
        void IDynamicProperty2.GetDisplayName(ref string szName)
        {
            szName = "My integer property";
        }

        // Show/Hide property in the OPM, for this object instance
        void IDynamicProperty2.IsPropertyEnabled(object pUnk, out int bEnabled)
        {
            bEnabled = 1;
        }

        // Is property showing but disabled
        void IDynamicProperty2.IsPropertyReadOnly(out int bReadonly)
        {
            bReadonly = 0;
        }

        // Get the property description string

        void IDynamicProperty2.GetDescription(ref string szName)
        {
            szName = "This property is an integer";
        }

        // OPM will typically display these in an edit field
        // optional: meta data representing property type name,
        // ex. ACAD_ANGLE
        void IDynamicProperty2.GetCurrentValueName(ref string szName)
        {
            throw new System.NotImplementedException();
        }

        // What is the property type, ex. VT_R8
        void IDynamicProperty2.GetCurrentValueType(out ushort varType)
        {
            // The Property Inspector supports the following data
            // types for dynamic properties:
            // VT_I2, VT_I4, VT_R4, VT_R8,VT_BSTR, VT_BOOL
            // and VT_USERDEFINED. 

            varType = 3; // VT_I4
        }

        // Get the property value, passes the specific object
        // we need the property value for.
        void IDynamicProperty2.GetCurrentValueData(object pUnk, ref object pVarData)
        {
            // TODO: Get the value and return it to AutoCAD
            int temp = 4;
            // Because we said the value type was a 32b int (VT_I4)

            // Convert the COM object to a managed object
            var id = DBObject.FromAcadObject(pUnk);

            using (Transaction tr = new OpenCloseTransaction())
            {
                Line l = !id.IsNull ? tr.GetObject(id, OpenMode.ForRead) as Line : null;

                tr.Abort();
            }

            

            pVarData = (int)4;
        }

        // Set the property value, passes the specific object we
        // want to set the property value for
        void IDynamicProperty2.SetCurrentValueData(object pUnk, object varData)
        {
            // TODO: Save the value returned to you

            // Because we said the value type was a 32b int (VT_I4)
            int myVal = (int)varData;
        }

        // OPM passes its implementation of IDynamicPropertyNotify, you
        // cache it and call it to inform OPM your property has changed
        void IDynamicProperty2.Connect(object pSink)
        {
            m_pSink = (IDynamicPropertyNotify2)pSink;
        }

        void IDynamicProperty2.Disconnect()
        {
            m_pSink = null;
        }

        public void MapPropertyToCategory(int dispid, ref int ppropcat)
        {
            ppropcat = 1;
        }

        public void GetCategoryName(int propcat, uint lcid, ref string pbstrName)
        {
            if (propcat != 1) pbstrName = string.Empty;
            pbstrName = "My Properties";
        }
    }



    #endregion

    #region Enum Prop

    #region Com Stuff

    [
        Guid("c71ff811-8b0d-4f64-ac6b-1b8fd206f71f"),
        ProgId("OPMNetSample.CustomEnumProperty.1"),

        // No class interface is generated for this class and
        // no interface is marked as the default.
        // Users are expected to expose functionality through
        // interfaces that will be explicitly exposed by the object
        // This means the object can only expose interfaces we define

        ClassInterface(ClassInterfaceType.None),
        // Set the default COM interface that will be used for
        // Automation. Languages like: C#, C++ and VB allow to 
        //query for interface's we're interested in but Automation 
        // only aware languages like javascript do not allow to 
        // query interface(s) and create only the default one

        ComDefaultInterface(typeof(IDynamicEnumProperty)),
        ComVisible(true)
    ]

    #endregion
    public class CustomEnumProp : IDynamicProperty2, ICategorizeProperties, IDynamicEnumProperty
    {
        #region IDynamicProperty2

        private IDynamicPropertyNotify2 m_pSink = null;

        // Unique property ID
        void IDynamicProperty2.GetGUID(out Guid propGUID)
        {
            propGUID = new Guid("c71ff811-8b0d-4f64-ac6b-1b8fd206f71f");
        }

        // Property display nameD
        void IDynamicProperty2.GetDisplayName(ref string szName)
        {
            szName = "My enum property";
        }

        // Show/Hide property in the OPM, for this object instance
        void IDynamicProperty2.IsPropertyEnabled(object pUnk, out int bEnabled)
        {
            bEnabled = 1;
        }

        // Is property showing but disabled
        void IDynamicProperty2.IsPropertyReadOnly(out int bReadonly)
        {
            bReadonly = 0;
        }

        // Get the property description string

        void IDynamicProperty2.GetDescription(ref string szName)
        {
            szName = "This property is an enum";
        }

        // OPM will typically display these in an edit field
        // optional: meta data representing property type name,
        // ex. ACAD_ANGLE
        void IDynamicProperty2.GetCurrentValueName(ref string szName)
        {
            szName = string.Empty;
        }

        // What is the property type, ex. VT_R8
        void IDynamicProperty2.GetCurrentValueType(out ushort varType)
        {
            // The Property Inspector supports the following data
            // types for dynamic properties:
            // VT_I2, VT_I4, VT_R4, VT_R8,VT_BSTR, VT_BOOL
            // and VT_USERDEFINED. 
            /*
               * VT_I2 => 2
               * VT_I4 => 3
               * VT_R4 => 4
               * VT_R8 => 5
               * VT_BSTR => 8
               * VT_BOOL => 11
               * VT_USERDEFINED => 29
             */

            varType = 8;
        }

        // Get the property value, passes the specific object
        // we need the property value for.
        void IDynamicProperty2.GetCurrentValueData(object pUnk, ref object pVarData)
        {
            // TODO: Get the value and return it to AutoCAD
            int temp = 1;
            // Because we said the value type was a 32b int (VT_I4)

            // Convert the COM object to a managed object
            //IntPtr pUnkIntPtr = Marshal.GetIUnknownForObject(pUnk);
            //Line line = (Line)Marshal.GetObjectForIUnknown(pUnkIntPtr);

            //var id = DBObject.FromAcadObject(pUnk);

            //using (Transaction tr = new OpenCloseTransaction())
            //{
            //    Line l = !id.IsNull ? tr.GetObject(id, OpenMode.ForRead) as Line : null;

            //    tr.Abort();
            //}



            pVarData = m_enumValues[1];
        }

        // Set the property value, passes the specific object we
        // want to set the property value for
        void IDynamicProperty2.SetCurrentValueData(object pUnk, object varData)
        {
            // TODO: Save the value returned to you

            // Because we said the value type was a 32b int (VT_I4)
            int myVal = (int)varData;
        }

        // OPM passes its implementation of IDynamicPropertyNotify, you
        // cache it and call it to inform OPM your property has changed
        void IDynamicProperty2.Connect(object pSink)
        {
            m_pSink = (IDynamicPropertyNotify2)pSink;
        }

        void IDynamicProperty2.Disconnect()
        {
            m_pSink = null;
        }

        #endregion

        #region IDynamicProperty

        //private IDynamicPropertyNotify m_pSink = null;

        //public void GetGUID(out Guid propGUID)
        //{
        //    propGUID = new Guid("c71ff811-8b0d-4f64-ac6b-1b8fd206f71f");
        //}

        //public void GetDisplayName(ref string bstrName)
        //{
        //    bstrName = "My enum property";
        //}

        //public void IsPropertyEnabled(IntPtr objectID, out bool pbEnabled)
        //{
        //    pbEnabled = true;
        //}

        //public void IsPropertyReadOnly(out bool pbReadonly)
        //{
        //    pbReadonly = false;
        //}

        //public void GetDescription(ref string bstrName)
        //{
        //    bstrName = "Please work";
        //}

        //public void GetCurrentValueName(ref string pbstrName)
        //{
        //    pbstrName = string.Empty;
        //}

        //public void GetCurrentValueType(out ushort pVarType)
        //{
        //    pVarType = 8;
        //}

        //public void GetCurrentValueData(IntPtr objectID, out object pvarData)
        //{
        //    pvarData = m_enumValues[1];
        //}

        //public void SetCurrentValueData(IntPtr objectID, object varData)
        //{

        //}

        //public void Connect(object pSink)
        //{
        //    m_pSink = (IDynamicPropertyNotify)pSink;
        //}

        //public void Disconnect()
        //{
        //    m_pSink = null;
        //}

        #endregion


        public void MapPropertyToCategory(int dispid, ref int ppropcat)
        {
            ppropcat = 1;
        }

        public void GetCategoryName(int propcat, uint lcid, ref string pbstrName)
        {
            if (propcat != 1) pbstrName = string.Empty;
            pbstrName = "My Properties";
        }



        private readonly string[] m_enumValues = { "Option 1", "Option 2", "Option 3" };

        void IDynamicEnumProperty.GetNumPropertyValues(ref int numValues)
        {
            numValues = m_enumValues.Length;
        }

        void IDynamicEnumProperty.GetPropValueName(int index, ref string valueName)
        {
            if (index >= 0 && index < m_enumValues.Length)
            {
                valueName = m_enumValues[index];
            }
            else
            {
                valueName = string.Empty;
            }
        }

        void IDynamicEnumProperty.GetPropValueData(int index, ref object valueData)
        {
            if (index >= 0 && index < m_enumValues.Length)
            {
                valueData = m_enumValues[index];
            }
            else
            {
                valueData = null;
            }
        }

        
    }

    #endregion

    #region Application Entry Point
    public class MyEntryPoint : IExtensionApplication {
        protected internal CustomProp custProp = null;
        protected internal CustomEnumProp enumProp = null;

        public void Initialize () {
            Assembly.LoadFrom(@"C:\Users\nshupe\source\repos\OPMNetExt\x64\Debug\asdkOPMNetExt.dll");

            // Add the Dynamic Property
            Dictionary classDict =SystemObjects.ClassDictionary;
            RXClass lineDesc =(RXClass)classDict.At("AcDbLine");
            IPropertyManager2 pPropMan =(IPropertyManager2)xOPM.xGET_OPMPROPERTY_MANAGER(lineDesc);

            custProp =new CustomProp();
            pPropMan.AddProperty((object)custProp);

            //IPropertyManager propMan = (IPropertyManager)xOPM.xGET_OPMPROPERTY_MANAGER(lineDesc);
            
            enumProp = new CustomEnumProp();
            pPropMan.AddProperty((object)enumProp);
        }

        public void Terminate () {
            // Remove the Dynamic Property
            Dictionary classDict =SystemObjects.ClassDictionary;
            RXClass lineDesc =(RXClass)classDict.At("AcDbLine");
            IPropertyManager2 pPropMan =(IPropertyManager2)xOPM.xGET_OPMPROPERTY_MANAGER(lineDesc);

            pPropMan.RemoveProperty((object)custProp);
            custProp =null;

            //IPropertyManager propMan = (IPropertyManager)xOPM.xGET_OPMPROPERTY_MANAGER(lineDesc);

            pPropMan.RemoveProperty((object)enumProp);
            enumProp = null;
        }

    }
    #endregion

}
