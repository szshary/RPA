using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace RPA.Core
{
    sealed internal class WordRuleSet : StatefulRuleSetDecorator
    {
        private SortedDictionary<String, ContentControl> documentTextContentControls;
        private Application wordApplication;
        private Document wordDocument;

        public WordRuleSet(StatefulRuleSet statefulRuleSet) : base(statefulRuleSet)
        {
            _elementStartRules.Add("FillPlainTextContentControl", FillPlainTextContentControl);
            _elementStartRules.Add("SaveDocumentAsNewFile", SaveDocumentAsNewFile);

            _elementStartRules.Add("WordSession", StartWordSession);
            _elementEndRules.Add("WordSession", EndWordSession);
        }

        private void FillPlainTextContentControl(Dictionary<String, String> parameters)
        {
            if (parameters.ContainsKey("Tag") && documentTextContentControls.ContainsKey(parameters["Tag"]))
            {
                if (parameters.ContainsKey("Value"))
                {
                    documentTextContentControls[parameters["Tag"]].Range.Text = parameters["Value"];
                }
                else if (parameters.ContainsKey("Variable") && EngineState.VariableCollection.ContainsKey(parameters["Variable"]))
                {
                    documentTextContentControls[parameters["Tag"]].Range.Text = EngineState.VariableCollection[parameters["Variable"]].ToString();
                }
            }
        }

        private void SaveDocumentAsNewFile(Dictionary<String, String> parameters)
        {
            if (parameters.ContainsKey("FileName"))
            {
                String fileName = parameters["FileName"];
                if (parameters.ContainsKey("SuffixVariable") && EngineState.VariableCollection.ContainsKey(parameters["SuffixVariable"]))
                {
                    fileName += EngineState.VariableCollection[parameters["SuffixVariable"]].ToString();
                }
                fileName += ".docx";
                Object useDefaultValue = Type.Missing;

                wordDocument.SaveAs2(Path.Combine(Directory.GetCurrentDirectory(), parameters["Folder"], fileName), useDefaultValue, useDefaultValue, useDefaultValue, useDefaultValue, useDefaultValue,
                                                                            useDefaultValue, useDefaultValue, useDefaultValue, useDefaultValue, useDefaultValue, useDefaultValue, useDefaultValue,
                                                                            useDefaultValue, useDefaultValue, useDefaultValue);
            }
        }

        private void StartWordSession(Dictionary<String, String> parameters)
        {
            Process[] orphanWordProcesses = Process.GetProcessesByName("WORD");
            foreach (Process orphanProcess in orphanWordProcesses)
            {
                //User Word process always have window name whereas COM process do not.
                if (orphanProcess.MainWindowTitle.Length == 0)
                {
                    orphanProcess.Kill();
                }
            }
            if (wordApplication == null)
            {
                wordApplication = new Application
                {
                    Visible = true,
                    WindowState = WdWindowState.wdWindowStateMaximize
                };
            }
            if (parameters.ContainsKey("TemplateFileName"))
            {
                if (File.Exists(Path.Combine(Directory.GetCurrentDirectory(), (parameters.ContainsKey("Folder") ? parameters["Folder"] : String.Empty), parameters["TemplateFileName"])))
                {
                    Object useDefaultValue = Type.Missing;

                    wordDocument = wordApplication.Documents.Open((Path.Combine(Directory.GetCurrentDirectory(), "WordDocuments", parameters["TemplateFileName"])),
                                                                                ref useDefaultValue, ref useDefaultValue, ref useDefaultValue, ref useDefaultValue, ref useDefaultValue, ref useDefaultValue,
                                                                                ref useDefaultValue, ref useDefaultValue, ref useDefaultValue, ref useDefaultValue, ref useDefaultValue, ref useDefaultValue,
                                                                                ref useDefaultValue, ref useDefaultValue, ref useDefaultValue);

                    documentTextContentControls = new SortedDictionary<string, ContentControl>();
                    foreach (ContentControl conCon in wordDocument.Content.ContentControls)
                    {
                        if (conCon.Type == WdContentControlType.wdContentControlText && conCon.Tag != null && !documentTextContentControls.ContainsKey(conCon.Tag))
                        {
                            documentTextContentControls.Add(conCon.Tag, conCon);
                        }
                    }
                    foreach (Microsoft.Office.Interop.Word.Shape shp in wordDocument.Shapes)
                    {
                        foreach (ContentControl conCon in shp.TextFrame.ContainingRange.ContentControls)
                        {
                            if (conCon.Type == WdContentControlType.wdContentControlText && conCon.Tag != null && !documentTextContentControls.ContainsKey(conCon.Tag))
                            {
                                documentTextContentControls.Add(conCon.Tag, conCon);
                            }
                        }
                    }
                }
            }
        }

        private void EndWordSession()
        {
            if (wordApplication != null)
            {
                if (wordDocument != null)
                {
                    wordDocument.Close();
                }
                wordApplication.WindowState = WdWindowState.wdWindowStateMinimize;
                wordApplication.Quit();

                Process[] orphanWordProcesses = Process.GetProcessesByName("WORD");
                foreach (Process orphanProcess in orphanWordProcesses)
                {
                    //User Word process always have window name whereas COM process do not.
                    if (orphanProcess.MainWindowTitle.Length == 0)
                    {
                        orphanProcess.Kill();
                    }
                }
            }
        }
    }
}