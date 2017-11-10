using System;
using System.Runtime.InteropServices;
using Inventor;
using Microsoft.Win32;
using InvAddIn;

namespace AssemblyToParts
{
    /// <summary>
    /// This is the primary AddIn Server class that implements the ApplicationAddInServer interface
    /// that all Inventor AddIns are required to implement. The communication between Inventor and
    /// the AddIn is via the methods on this interface.
    /// </summary>
    [GuidAttribute("124e7311-6a45-4301-8485-29fc60060a1f")]
    public class StandardAddInServer : Inventor.ApplicationAddInServer
    {

        // Inventor application object.
        private Inventor.Application m_inventorApplication;
        private ButtonDefinition m_PushParametersButton;
        private ButtonDefinition m_UpdateIlogicButtton;

        private static string addInGUID = "124e7311-6a45-4301-8485-29fc60060a1f";

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

            // TODO: Add ApplicationAddInServer.Activate implementation.
            // e.g. event initialization, command creation etc.

            ControlDefinitions controlDefs = m_inventorApplication.CommandManager.ControlDefinitions;

            m_PushParametersButton = controlDefs.AddButtonDefinition("Push\nParameters", "PushParameters", CommandTypesEnum.kShapeEditCmdType, addInGUID, "Push Assembly Parameters to child parts.", "Push Parameters");
            m_UpdateIlogicButtton = controlDefs.AddButtonDefinition("Update\niLogic", "UpdateIlogic", CommandTypesEnum.kShapeEditCmdType, addInGUID, "Update Ilogic for passing assembly parameters to children.", "Update iLogic");

            if (firstTime)
            {

                try
                {
                    if (m_inventorApplication.UserInterfaceManager.InterfaceStyle == InterfaceStyleEnum.kRibbonInterface)
                    {
                        Ribbon ribbon = m_inventorApplication.UserInterfaceManager.Ribbons["Assembly"];
                        RibbonTab tab = ribbon.RibbonTabs["id_TabAssemble"];

                        try
                        {
                            // For ribbon interface
                            // This is a new panel that can be made
                            RibbonPanel panel = tab.RibbonPanels.Add("Assembly to Parts", "Autodesk:Assembly to Parts:Panel1", addInGUID, "", false);
                            CommandControl control1 = panel.CommandControls.AddButton(m_PushParametersButton, true, true, "", false);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                    }
                    else
                    {
                        // For classic interface, possibly incorrect code
                        CommandBar oCommandBar = m_inventorApplication.UserInterfaceManager.CommandBars["AMxAssemblyPanelCmdBar"];
                        oCommandBar.Controls.AddButton(m_PushParametersButton, 0);
                        oCommandBar.Controls.AddButton(m_UpdateIlogicButtton, 0);
                    }
                }
                catch
                {
                    // For classic interface, possibly incorrect code
                    CommandBar oCommandBar = m_inventorApplication.UserInterfaceManager.CommandBars["AMxAssemblyPanelCmdBar"];
                    oCommandBar.Controls.AddButton(m_PushParametersButton, 0);
                    oCommandBar.Controls.AddButton(m_UpdateIlogicButtton, 0);
                }
            }

            m_PushParametersButton.OnExecute += new ButtonDefinitionSink_OnExecuteEventHandler(m_PushParametersButton_OnExecute);
            m_UpdateIlogicButtton.OnExecute += new ButtonDefinitionSink_OnExecuteEventHandler(m_UpdateIlogicButtton_OnExecute);
        }

        public void Deactivate()
        {
            // This method is called by Inventor when the AddIn is unloaded.
            // The AddIn will be unloaded either manually by the user or
            // when the Inventor session is terminated

            // TODO: Add ApplicationAddInServer.Deactivate implementation

            // Release objects.
            m_inventorApplication = null;

            Marshal.ReleaseComObject(m_PushParametersButton);
            m_PushParametersButton = null;

            Marshal.ReleaseComObject(m_UpdateIlogicButtton);
            m_UpdateIlogicButtton = null;

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

        public void m_PushParametersButton_OnExecute(NameValueMap Context)
        {
            ParameterList parameterList = new ParameterList();
            parameterList.pushChildren();
        }
        public void m_UpdateIlogicButtton_OnExecute(NameValueMap Context)
        {
           
        }


        #endregion

    }
}
