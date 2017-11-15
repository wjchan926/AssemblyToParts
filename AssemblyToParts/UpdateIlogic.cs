using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Inventor;
using System.Collections;

namespace InvAddIn
{
    class UpdateIlogic
    {
        // ilogic is a prexisting addin
        private static string iLogicAddinGuid = "{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}";

        private Application invApp;
        private ApplicationAddIn iLogicAddin;
        private Document currentAssembly;
        private String ruleName = "Dimensions to Parts";

        public UpdateIlogic()
        {
            invApp = (Application)Marshal.GetActiveObject("Inventor.Application");
            currentAssembly = invApp.ActiveDocument;
            iLogicAddin = invApp.ApplicationAddIns.ItemById[iLogicAddinGuid];           
        }

        public void updateParametersRule()
        {
            if (iLogicAddin != null)            {

                // Get Key Parameters Using custom class
                ArrayList keyParameters = new ArrayList();
                UserParameters userParams = ((AssemblyDocument)currentAssembly).ComponentDefinition.Parameters.UserParameters;

                // Add all key parameters to list
                foreach (UserParameter userParam in userParams)
                {
                    if (userParam.IsKey)
                    {
                        keyParameters.Add(userParam);
                    }
                }

                // Create iLogic code as string
                StringBuilder iLogicCode = new StringBuilder();
                iLogicCode.Append("Dim oDoc As Document\noDoc = ThisDoc.Document\nDim partFile As Document\n");
                iLogicCode.AppendLine("For Each partFile In oDoc.AllReferencedDocuments");
                iLogicCode.AppendLine("Dim fileNameLocation As Long");
                iLogicCode.AppendLine("fileNameLocation = InStrRev(partFile.FullFileName, \"\\\", -1)");
                iLogicCode.Append("Dim oPart As String\noPart = Right(partFile.FullFileName, Len(partFile.FullFileName) - fileNameLocation)\n");
                iLogicCode.Append("Parameter.Quiet = True\nTry\n");

                // Add Parameters to iLogic code
                foreach (UserParameter userParam in keyParameters)
                {
                    iLogicCode.AppendLine(parameterParser(userParam.Name));
                }

                iLogicCode.Append("Catch\nEnd Try\nInventorVb.DocumentUpdate()\nNext\n");

                // Enter Addin
                dynamic iLogicAuto = iLogicAddin.Automation;

                //Update Code
                try
                {
                    iLogicAuto.DeleteRule(currentAssembly, ruleName);
                } catch(Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(ex.ToString());
                }
                finally
                {
                    iLogicAuto.AddRule(currentAssembly, ruleName, iLogicCode.ToString());
                }
            }       
        }

        private string parameterParser(string parameter)
        {            
            return "Parameter(oPart, \"" + parameter + "\") = " + parameter;
        }
    }
}
