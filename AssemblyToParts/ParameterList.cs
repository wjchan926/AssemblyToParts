using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;
using System.Runtime.InteropServices;

namespace InvAddIn
{
    class ParameterList
    {
        private Application invApp;
        private UserParameters userParameters;
        private AssemblyDocument currentAssembly;

        public ParameterList()
        {
            invApp = (Inventor.Application)Marshal.GetActiveObject("Inventor.Application");
            currentAssembly = (AssemblyDocument)invApp.ActiveDocument;
            userParameters = currentAssembly.ComponentDefinition.Parameters.UserParameters;
        }

        public void pushChildren()
        {
            foreach (PartDocument partFile in userParameters)
            {
                try
                {
                    UserParameters partParameters = partFile.ComponentDefinition.Parameters.UserParameters;

                    foreach (UserParameter userParameter in userParameters)
                    {
                        if (userParameter.IsKey)
                        {
                            partParameters.AddByExpression(userParameter.Name, userParameter.Expression, userParameter.get_Units());
                        }
                    }

                }
                catch (Exception ex)
                {

                }
            }
        }

    }
}
