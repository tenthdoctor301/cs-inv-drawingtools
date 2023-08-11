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
        private Inventor.Application invApp;
        private PartDetails partDetails;
        

        public ButtonDefinitionSink_OnExecuteEventHandler ButtonDefinition_OnExecuteEventDelegate;
        public StandardAddInServer()
        {
        }

        #region ApplicationAddInServer Members
        public void Activate(Inventor.ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            invApp = addInSiteObject.Application;
            partDetails = new PartDetails();
            partDetails.OnExecute(invApp, AddInClientID());
            // Add to the user interface, if it's the first time.
            
            if (firstTime == true)
            {

                partDetails.AddToUserInterface(AddInClientID());
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
            invApp = null;

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
       // private invTools.PartDetails m_PartDetails;
       




        public string AddInClientID()
        {
            GuidAttribute addInCLSID;
            addInCLSID = (GuidAttribute)GuidAttribute.GetCustomAttribute(typeof(StandardAddInServer), typeof(GuidAttribute));
            string addInCLSIDString;
            addInCLSIDString = "{" + addInCLSID.Value + "}";
            return addInCLSIDString;
        }


        #endregion

    }
}
