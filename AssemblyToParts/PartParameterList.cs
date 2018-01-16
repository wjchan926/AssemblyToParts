using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;
using System.Runtime.InteropServices;
using System.Collections;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace InvAddIn
{
    class PartParameterList
    {

        private Inventor.Application invApp;
        private UserParameters assemblyParameters;
        private AssemblyComponentDefinition parentAssembly;
        private AssemblyDocument assemblyDoc;
        private PartDocument partDoc;
        private PartComponentDefinition currentPartDef;

        public PartParameterList(){
            invApp = (Inventor.Application)Marshal.GetActiveObject("Inventor.Application");
            
            // Must Detect if in In-Place Edit Session
            partDoc = (PartDocument)invApp.ActiveEditDocument;
            currentPartDef = partDoc.ComponentDefinition;
        }

        public void openPullParents()
        {
            // Added the First Generation Parent as default for the parent V3.1.0
            FileBrowser fileBrowser = new FileBrowser(openFolder(), invApp.ActiveDocument.FullFileName);
            fileBrowser.ShowDialog();
         
            if (fileBrowser.DialogResult == DialogResult.OK)
            {
                string parentFile = fileBrowser.getFileChosen();

                pullParents(parentFile);

            }
        }

        private string openFolder()
        {
            string folderPath = null;

            try
            {
                folderPath = partDoc.File.FullFileName;
                int lastIndex = folderPath.LastIndexOf("\\");
                folderPath = folderPath.Substring(0, lastIndex + 1);
            }
            catch (Exception)
            {
                MessageBox.Show("Part File Has Not Been Saved.", "File Not Saved");
                folderPath = @"C:\_Vault\Designs";
            }
            
            return folderPath;

        }

        private void pullParents(string parentFileName)
        {
            // Set parent assembly do to parentFileName
            try
            {
                invApp.Documents.Open(parentFileName, false);
            }
            catch (Exception e)
            {
                throw;
            }
            finally
            {
                assemblyDoc = (AssemblyDocument)invApp.Documents.ItemByName[parentFileName];
                parentAssembly = assemblyDoc.ComponentDefinition;
                assemblyParameters = parentAssembly.Parameters.UserParameters;
            }
            // Get User Parameters from parent file.
            ArrayList parentParameters = new ArrayList();

            foreach (Parameter p in assemblyParameters)
            {
                // Store parameter in list if key
                if (p.IsKey)
                {
                    parentParameters.Add(p);
                }
            }

            // Add all parameters from list to part file
            // Check if parameter already exists

            // Add all part parameters to list for comparision
            ArrayList partParameters = new ArrayList();
            UserParameters childParameters = currentPartDef.Parameters.UserParameters;
            foreach (Parameter p in childParameters)
            {
                partParameters.Add(p.Name);
            }

            // Add parent parameters to part parameters
            foreach (Parameter p in parentParameters)
            {
                //Check if parameter already exists
                if (partParameters.Contains(p.Name))
                {
                    // Try to replace value with parent value
                    try
                    {
                        childParameters[p.Name].Value = p._Value;
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show("Error:\n" + e.Message);
                    }
                }
                else
                {
                    // Try to add the parameter by nominal value
                    try
                    {
                        childParameters.AddByValue(p.Name, p._Value, p.get_Units());
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show("Error:\n" + e.Message);
                    }
                }
            }


        }

        /**
         * Need to open invisible and get data
         **/
        private void setParentAssembly(String parentFullName)
        {
            try
            {
                assemblyDoc = (AssemblyDocument)invApp.Documents.Open(parentFullName, false);
            }
            catch(Exception ex)
            {
                MessageBox.Show("Document already open.\nError:\n" + ex.Message);
                
            }
            finally
            {
                
            }  
           
        }

    }

}
