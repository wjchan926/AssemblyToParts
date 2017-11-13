using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace InvAddIn
{
    class UpdateIlogic
    {
        // ilogic is a prexisting addin
        private static string iLogicAddinGuid = "{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}";

        private Inventor.Application invApp;
        private Inventor.ApplicationAddIn iLogicAddin;
        private Inventor.AssemblyDocument currentAssembly;

        public UpdateIlogic()
        {
            invApp = (Inventor.Application)Marshal.GetActiveObject("Inventor.Application");
            currentAssembly = (Inventor.AssemblyDocument)invApp.ActiveDocument;
            iLogicAddin = invApp.ApplicationAddIns.ItemById[iLogicAddinGuid];           
        }

        public void updateParametersRule()
        {
             
      //      iLogicAddin.Automation.GetRule(currentAssembly, "Dimensions to Parts");
              
        }


    }
}
