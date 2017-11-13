using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;
using System.Runtime.InteropServices;
using System.Collections;

namespace InvAddIn
{
    class ParameterList
    {
        private Application invApp;
        private UserParameters assemblyParameters;
        private AssemblyComponentDefinition currentAssembly;
        private AssemblyDocument assemblyDoc;
        

        public ParameterList()
        {
            invApp = (Application)Marshal.GetActiveObject("Inventor.Application");
            assemblyDoc = (AssemblyDocument)invApp.ActiveDocument;
            currentAssembly = assemblyDoc.ComponentDefinition;
            assemblyParameters = currentAssembly.Parameters.UserParameters;
        }

        public void pushChildren()
        {
            DocumentsEnumerator allRefDocs = assemblyDoc.AllReferencedDocuments;

            try
            {
                foreach (Document oDoc in allRefDocs)
                {
                    if (oDoc.DocumentType.Equals(DocumentTypeEnum.kPartDocumentObject))
                    {
                        // For Parts
                        PartDocument partFile = (PartDocument)oDoc;
                        PartComponentDefinition partFileDef = partFile.ComponentDefinition;
                        UserParameters partParameters = partFileDef.Parameters.UserParameters;
                        ArrayList paraList = new ArrayList();

                        foreach (Parameter parameter in partParameters)
                        {
                            paraList.Add(parameter.Name);
                        }

                        foreach (Parameter parameter in assemblyParameters)
                        {
                            if (parameter.IsKey && !paraList.Contains(parameter.Name))
                            {
                                partParameters.AddByExpression(parameter.Name, parameter.Expression, parameter.get_Units());
                            }
                        }
                    }

                    else if (oDoc.DocumentType.Equals(DocumentTypeEnum.kAssemblyDocumentObject))
                    {
                        // For Subassemblies
                        AssemblyDocument partFile = (AssemblyDocument)oDoc;
                        AssemblyComponentDefinition subAssemblyDef = partFile.ComponentDefinition;
                        UserParameters subAssemblyParameters = subAssemblyDef.Parameters.UserParameters;

                        ArrayList paraList = new ArrayList();

                        foreach (Parameter parameter in subAssemblyParameters)
                        {
                            paraList.Add(parameter.Name);
                        }

                        foreach (Parameter parameter in assemblyParameters)
                        {
                            if (parameter.IsKey && !paraList.Contains(parameter.Name))
                            {
                                subAssemblyParameters.AddByExpression(parameter.Name, parameter.Expression, parameter.get_Units());
                            }
                        }
                    }

                }
            } catch (Exception ex)
            {

            }          
        }

    }
}
