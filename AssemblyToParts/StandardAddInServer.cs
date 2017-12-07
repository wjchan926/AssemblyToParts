using System;
using System.Runtime.InteropServices;
using Inventor;
using Microsoft.Win32;
using InvAddIn;
using System.Drawing;
using System.Windows.Forms;


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
        private ButtonDefinition m_PushAndUpdateButton;

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

            // Icons  
            //   System.Diagnostics.Debug.WriteLine(System.IO.Directory.GetCurrentDirectory());

            //string originalDir = System.IO.Directory.GetCurrentDirectory();
            //string dir = "%APPDATA%\\Autodesk\\ApplicationPlugins\\AssemblyToParts\\";
            //System.Diagnostics.Debug.WriteLine(System.IO.Directory.GetCurrentDirectory());
            //System.IO.Directory.SetCurrentDirectory(dir);
            //MessageBox.Show(System.IO.Directory.GetCurrentDirectory());

            string appData = System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData);

            Icon smallPush = new Icon(appData + @"\Autodesk\ApplicationPlugins\AssemblyToParts\push parameters.ico");
            Icon largePush = new Icon(appData + @"\Autodesk\ApplicationPlugins\AssemblyToParts\\push parameters.ico");

        //    System.IO.Directory.SetCurrentDirectory(originalDir);

            stdole.IPictureDisp smallPushIcon = PictureDispConverter.ToIPictureDisp(smallPush);
            stdole.IPictureDisp largePushIcon = PictureDispConverter.ToIPictureDisp(largePush);
            // End Icon code
            
            m_PushAndUpdateButton = controlDefs.AddButtonDefinition("Push and\nUpdate", "UpdateIlogic", CommandTypesEnum.kShapeEditCmdType, addInGUID, "Push Assembly Parameters to child parts and/nupdate Ilogic for passing assembly parameters to children.",
                "Push Parameters and Update iLogic", smallPushIcon, largePushIcon);

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
                            //   CommandControl control1 = panel.CommandControls.AddButton(m_PushParametersButton, true, true, "", false);
                            //  CommandControl control2 = panel.CommandControls.AddButton(m_UpdateIlogicButtton, true, true, "", false);
                            CommandControl control1 = panel.CommandControls.AddButton(m_PushAndUpdateButton, true, true, "", false);
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
                        oCommandBar.Controls.AddButton(m_PushAndUpdateButton, 0);
                    //    oCommandBar.Controls.AddButton(m_PushParametersButton, 0);
                    //oCommandBar.Controls.AddButton(m_UpdateIlogicButtton, 0);
                    }
                }
                catch
                {
                    // For classic interface, possibly incorrect code
                    CommandBar oCommandBar = m_inventorApplication.UserInterfaceManager.CommandBars["AMxAssemblyPanelCmdBar"];
                    oCommandBar.Controls.AddButton(m_PushAndUpdateButton, 0);
                    //    oCommandBar.Controls.AddButton(m_PushParametersButton, 0);
                    //oCommandBar.Controls.AddButton(m_UpdateIlogicButtton, 0);
                }
            }

            //  m_PushParametersButton.OnExecute += new ButtonDefinitionSink_OnExecuteEventHandler(m_PushParametersButton_OnExecute);
            //  m_UpdateIlogicButtton.OnExecute += new ButtonDefinitionSink_OnExecuteEventHandler(m_UpdateIlogicButtton_OnExecute);
            m_PushAndUpdateButton.OnExecute += new ButtonDefinitionSink_OnExecuteEventHandler(m_PushAndUpdateButton_OnExecute);
        }

        public void Deactivate()
        {
            // This method is called by Inventor when the AddIn is unloaded.
            // The AddIn will be unloaded either manually by the user or
            // when the Inventor session is terminated

            // TODO: Add ApplicationAddInServer.Deactivate implementation

            // Release objects.
            m_inventorApplication = null;

            /*
            Marshal.ReleaseComObject(m_PushParametersButton);
            m_PushParametersButton = null;

            Marshal.ReleaseComObject(m_UpdateIlogicButton);
            m_UpdateIlogicButton = null;*/

            Marshal.ReleaseComObject(m_PushAndUpdateButton);
            m_PushAndUpdateButton = null;

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
        /*
        public void m_PushParametersButton_OnExecute(NameValueMap Context)
        {
            ParameterList parameterList = new ParameterList();
            parameterList.pushChildren();
        }

        public void m_UpdateIlogicButtton_OnExecute(NameValueMap Context)
        {
            UpdateIlogic updateIlogic = new UpdateIlogic();
            updateIlogic.updateParametersRule();
        }*/

        public void m_PushAndUpdateButton_OnExecute(NameValueMap Context)
        {
            ParameterList parameterList = new ParameterList();
            parameterList.pushChildren();

            UpdateIlogic updateIlogic = new UpdateIlogic();
            updateIlogic.updateParametersRule();
        }



        #endregion
        
    }

    public sealed class PictureDispConverter
    {
        [DllImport("OleAut32.dll", EntryPoint = "OleCreatePictureIndirect", ExactSpelling = true, PreserveSig = false)]
        private static extern stdole.IPictureDisp OleCreatePictureIndirect([MarshalAs(UnmanagedType.AsAny)]
            object picdesc, ref Guid iid, [MarshalAs(UnmanagedType.Bool)]
            bool fOwn);


        static Guid iPictureDispGuid = typeof(stdole.IPictureDisp).GUID;

        private sealed class PICTDESC
        {
            private PICTDESC()
            {
            }


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


        public static stdole.IPictureDisp ToIPictureDisp(System.Drawing.Icon icon)
        {
            PICTDESC.Icon pictIcon = new PICTDESC.Icon(icon);
            return OleCreatePictureIndirect(pictIcon, ref iPictureDispGuid, true);
        }


        public static stdole.IPictureDisp ToIPictureDisp(System.Drawing.Bitmap bmp)
        {
            PICTDESC.Bitmap pictBmp = new PICTDESC.Bitmap(bmp);
            return OleCreatePictureIndirect(pictBmp, ref iPictureDispGuid, true);
        }
    }
}
