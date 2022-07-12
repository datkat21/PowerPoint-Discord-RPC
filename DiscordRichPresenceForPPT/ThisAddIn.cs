using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using DiscordRPC;
using DiscordRPC.Logging;
using System.Diagnostics;

namespace DiscordRichPresenceForPPT
{
    public partial class ThisAddIn
    {
        public DiscordRpcClient Rpc;
        public string Detail = "Idle";
        public string State = "Not doing anything";

        public void UpdatePresence()
        {
            Rpc.SetPresence(new RichPresence()
            {
                Details = Detail,
                State = State,
                Assets = new Assets()
                {
                    LargeImageKey = "logo",
                    LargeImageText = "Microsoft PowerPoint | DiscordRichPresenceForPPT v1.0",
                }
            });
        }

        public string TrimString(string Str, int Length = 50)
        {
            if (Str.Length > Length)
            {
                return Str.Substring(0, Length);
            }
            else
            {
                return Str;
            }
        }

        public string GetState(PowerPoint.Selection Sel)
        {
            string RetState = "";
            try
            {
                if (Sel.Type == PowerPoint.PpSelectionType.ppSelectionNone && Sel.SlideRange != null)
                {
                    RetState = $"Idle on Slide {Sel.SlideRange[1].SlideNumber}";
                }
                if (Sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
                {
                    RetState = $"Typing into {TrimString(Sel.ShapeRange[1].Name)} on Slide {Sel.SlideRange[1].SlideNumber}";
                }
                else if (Sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    if (Sel.ShapeRange.Count > 1)
                    {
                        RetState = $"Selecting multiple shapes on Slide {Sel.SlideRange[1].SlideNumber}";
                    }
                    else
                    {
                        RetState = $"Selecting {TrimString(Sel.ShapeRange[1].Name)} on Slide {Sel.SlideRange[1].SlideNumber}";
                    }
                }
                else if (Sel.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
                {
                    if (Sel.SlideRange.Count > 1)
                    {
                        RetState = "Selecting multiple slides";
                    }
                    else
                    {
                        RetState = $"Selecting Slide {Sel.SlideRange[1].SlideNumber}";
                    }
                }
            }
            catch (Exception)
            {
                Detail = $"Editing \"{TrimString(Application.ActiveWindow.Caption)}\"";
                RetState = "Idling";
            }
            return RetState;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Globals.ThisAddIn.Application.WindowSelectionChange += WindowSelection;
            Application.WindowActivate += WindowActivate;
            Application.PresentationClose += PresentationClose;

            Rpc = new DiscordRpcClient("996513202613518428");
            Rpc.Logger = new ConsoleLogger() { Level = LogLevel.Warning };
            Rpc.Initialize();
        }

        private void PresentationClose(PowerPoint.Presentation Pres)
        {
            Detail = "Idle";
            State = "";
            UpdatePresence();
        }

        private void WindowActivate(PowerPoint.Presentation Pres, PowerPoint.DocumentWindow Wn)
        {
            Detail = $"Editing \"{TrimString(Wn.Caption)}\"";
            Debug.Print($"Detail update: {Detail}");
            string TempState = GetState(Wn.Selection);

            if (TempState != State)
            {
                State = TempState;
                Debug.Print($"  State update: {State}");
            }

            UpdatePresence();
        }

        private void WindowSelection(PowerPoint.Selection Sel)
        {
            string TempState = GetState(Sel);

            if (State != TempState)
            {
                State = TempState;
                Debug.Print("State changed: " + State);
                UpdatePresence();
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Rpc.Dispose();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
