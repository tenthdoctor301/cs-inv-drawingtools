using Inventor;
using Microsoft.Win32;
using System;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Windows.Forms;


namespace DrawingTools
{
    
    /// <summary>
    /// This is the primary AddIn Server class that implements the ApplicationAddInServer interface
    /// that all Inventor AddIns are required to implement. The communication between Inventor and
    /// the AddIn is via the methods on this interface.
    /// </summary>
    [GuidAttribute("4043c68e-6202-48ff-a26b-24f230faab9b")]
    public class StandardAddInServer : Inventor.ApplicationAddInServer
    {

        // Inventor application object.
        Inventor.Application m_inventorApplication;
        ButtonDefinition CreateDetailsButton;
        //UserInterfaceEvents m_uiEvents;
        //ApplicationEvents m_appEvents;
        private ButtonDefinitionSink_OnExecuteEventHandler ButtonDefinition_OnExecuteEventDelegate;
        public StandardAddInServer()
        {
        }

        #region ApplicationAddInServer Members
        public void Activate(Inventor.ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            // This method is called by Inventor when it loads the addin.
            // The AddInSiteObject provides access to the Inventor Application object.
            // The FirstTime flag indicates if the addin is loaded for the first time.

            // Initialize AddIn members.
            m_inventorApplication = addInSiteObject.Application;
            //m_uiEvents = m_inventorApplication.UserInterfaceManager.UserInterfaceEvents;
            //m_appEvents = m_inventorApplication.ApplicationEvents;
            
            // creating a button definition.
            stdole.IPictureDisp largeIcon = PictureDispConverter.ToIPictureDisp(InvAddIn.Properties.Resources._32x32);
            stdole.IPictureDisp smallIcon = PictureDispConverter.ToIPictureDisp(InvAddIn.Properties.Resources._16x16);
            ControlDefinitions controlDefs = m_inventorApplication.CommandManager.ControlDefinitions;
            CreateDetailsButton = controlDefs.AddButtonDefinition("View Details", "id_viewdetails_bt", CommandTypesEnum.kShapeEditCmdType, AddInClientID(), "", "", smallIcon, largeIcon);
            ButtonDefinition_OnExecuteEventDelegate = new ButtonDefinitionSink_OnExecuteEventHandler(this.CreateDetailsButton_OnExecute);
            CreateDetailsButton.OnExecute += ButtonDefinition_OnExecuteEventDelegate;

            // Add to the user interface, if it's the first time.
            if (firstTime == true)
            {
                AddToUserInterface();
            }
            
            // TODO: Add ApplicationAddInServer.Activate implementation.
            // e.g. event initialization, command creation etc.
        }
        
        public void Deactivate()
        {
            // This method is called by Inventor when the AddIn is unloaded.
            // The AddIn will be unloaded either manually by the user or
            // when the Inventor session is terminated

            // TODO: Add ApplicationAddInServer.Deactivate implementation

            // Release objects.
            m_inventorApplication = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public void ExecuteCommand(int commandID)
        {
            // Note:this method is now obsolete, you should use the 
            // ControlDefinition functionality for implementing commands.
        }

        public object Automation
        {
            // This property is provided to allow the AddIn to expose an API 
            // of its own to other programs. Typically, this  would be done by
            // implementing the AddIn's API interface in a class and returning 
            // that class object through this property.

            get
            {
                // TODO: Add ApplicationAddInServer.Automation getter implementation
                return null;
            }
        }
       
        public void CreateDetailsButton_OnExecute(NameValueMap context)
        {
            invTools.PartDetails r = new invTools.PartDetails();
            r.ViewDetails();
            
        }

        private void AddToUserInterface()
        {
            RibbonPanel panel;
            try
            {
                panel = m_inventorApplication.UserInterfaceManager.Ribbons["Drawing"].RibbonTabs["id_TabPlaceViews"].RibbonPanels["id_UserToolTab"];
            }
            catch
            {
                panel = m_inventorApplication.UserInterfaceManager.Ribbons["Drawing"].RibbonTabs["id_TabPlaceViews"].RibbonPanels.Add("Tool Panel", "id_UserToolTab", AddInClientID(), "",false);
            }
            panel.CommandControls.AddButton(CreateDetailsButton);
        }

        public string AddInClientID()
        {
            GuidAttribute addInCLSID;
            addInCLSID = (GuidAttribute)GuidAttribute.GetCustomAttribute(typeof(StandardAddInServer), typeof(GuidAttribute));
            string addInCLSIDString;
            addInCLSIDString = "{" + addInCLSID.Value + "}";
            return addInCLSIDString;
        }

        public sealed class PictureDispConverter
        {
            [DllImport("OleAut32.dll", EntryPoint = "OleCreatePictureIndirect", ExactSpelling = true, PreserveSig = false)]
            private static extern stdole.IPictureDisp
                OleCreatePictureIndirect([MarshalAs(UnmanagedType.AsAny)] object picdesc,ref Guid iid,[MarshalAs(UnmanagedType.Bool)] bool fOwn);
            static Guid iPictureDispGuid = typeof(stdole.IPictureDisp).GUID;
            private static class PICTDESC
            {
                //Picture Types
                public const short PICTYPE_UNINITIALIZED = -1;
                public const short PICTYPE_NONE = 0;
                public const short PICTYPE_BITMAP = 1;
                public const short PICTYPE_METAFILE = 2;
                public const short PICTYPE_ICON = 3;
                public const short PICTYPE_ENHMETAFILE = 4;
                [StructLayout(LayoutKind.Sequential)]
                public class Icon
                {
                    internal int cbSizeOfStruct = Marshal.SizeOf(typeof(PICTDESC.Icon));
                    internal int picType = PICTDESC.PICTYPE_ICON;
                    internal IntPtr hicon = IntPtr.Zero;
                    internal int unused1;
                    internal int unused2;
                    internal Icon(System.Drawing.Icon icon)
                    {
                        this.hicon = icon.ToBitmap().GetHicon();
                    }
                }
                [StructLayout(LayoutKind.Sequential)]
                public class Bitmap
                {
                    internal int cbSizeOfStruct = Marshal.SizeOf(typeof(PICTDESC.Bitmap));
                    internal int picType = PICTDESC.PICTYPE_BITMAP;
                    internal IntPtr hbitmap = IntPtr.Zero;
                    internal IntPtr hpal = IntPtr.Zero;
                    internal int unused;
                    internal Bitmap(System.Drawing.Bitmap bitmap)
                    {
                        this.hbitmap = bitmap.GetHbitmap();
                    }
                }
            }

            public static stdole.IPictureDisp ToIPictureDisp(
                System.Drawing.Icon icon)
            {
                PICTDESC.Icon pictIcon = new PICTDESC.Icon(icon);
                return OleCreatePictureIndirect(
                    pictIcon, ref iPictureDispGuid, true);
            }
            public static stdole.IPictureDisp ToIPictureDisp(
                System.Drawing.Bitmap bmp)
            {
                PICTDESC.Bitmap pictBmp = new PICTDESC.Bitmap(bmp);
                return OleCreatePictureIndirect(pictBmp, ref iPictureDispGuid, true);
            }
        }
        #endregion

    }
}
