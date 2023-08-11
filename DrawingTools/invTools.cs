using Inventor;
using Microsoft.Win32;
using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Drawing;

namespace DrawingTools
{

    class PartDetails
    {

        public ButtonDefinitionSink_OnExecuteEventHandler ButtonDefinition_OnExecuteEventDelegate;
        ButtonDefinition CreateDetailsButton;
        private Inventor.Application invApp;

        public void OnExecute(Inventor.Application invApp, string AddInClientID)
        {
            this.invApp = invApp;
            Document oDoc = invApp.ActiveDocument;
            stdole.IPictureDisp largeIcon = PictureDispConverter.ToIPictureDisp(InvAddIn.Properties.Resources._32x32);
            stdole.IPictureDisp smallIcon = PictureDispConverter.ToIPictureDisp(InvAddIn.Properties.Resources._16x16);
            ControlDefinitions controlDefs = invApp.CommandManager.ControlDefinitions;
            CreateDetailsButton = controlDefs.AddButtonDefinition("View Details", "id_viewdetails_bt", CommandTypesEnum.kShapeEditCmdType, AddInClientID, "", "", smallIcon, largeIcon);
            ButtonDefinition_OnExecuteEventDelegate = new ButtonDefinitionSink_OnExecuteEventHandler(this.ViewDetails);
            CreateDetailsButton.OnExecute += ButtonDefinition_OnExecuteEventDelegate;

        }

        public void ViewDetails(NameValueMap context)
        {
            
            
            Document invDoc = invApp.ActiveDocument;
            
            Point2d oPosition = null;
           
            //Check if none selected
            SelectSet oS = invDoc.SelectSet;
            
            if (oS.Count == 0)
            {
                MessageBox.Show("Select a drawing view");
                return;
            }
            
            //Reference to the drawing view from the 1st selected object
            DrawingView oView = (DrawingView)oS[1];
            Inventor.Document viewDoc = (Document)oView.ReferencedDocumentDescriptor.ReferencedDocument;
            Property propTitle, propComments, propKeywords, propQty;
            String valueTitle, valueKeywords, valueComments, valueQtyFinal;
            int valueQty;
            propTitle = viewDoc.PropertySets["Summary Information"]["Title"];
            
            propComments = viewDoc.PropertySets["Summary Information"]["Comments"];
            propKeywords = viewDoc.PropertySets["Summary Information"]["Keywords"];
            try
            {
                propQty = viewDoc.PropertySets["User Defined Properties"]["Quantity"];
                valueQty = (int)propQty.Value;
            }
            catch
            {
                valueQty = 0;
            }

            valueTitle = (String)propTitle.Value;
            valueKeywords = (String)propKeywords.Value;
            valueComments = (String)propComments.Value;

            // Check if Comments is empty then Comments = null
            if (valueComments == "")
            {
                valueComments = "null";
            }

            // Check if Qty < 10 Then add "0" to first
            if (valueQty < 10)
            {
                valueQtyFinal = "0" + Convert.ToString(valueQty);
            }
            else
            {
                valueQtyFinal = Convert.ToString(valueQty);
            }

            // Remove space in head text
            if (Left(valueTitle, 1) == " ")
            {
                valueTitle = Right(valueTitle, valueTitle.Length - 1);
            
            }

            // Content
            String sFormattedText;
            sFormattedText = "<StyleOverride FontSize = '0.28'><StyleOverride FontSize = '0.305' Underline = 'True'>" + "Chi tiết: " + valueTitle + "</StyleOverride>" +
                "\r\n" + "Ký hiệu: " + valueKeywords + "\r\n" + "Số lượng: " + valueQtyFinal + "\r\n" + "Vật liệu: " + valueComments + "</StyleOverride>";
            DrawingDocument drawDoc = (DrawingDocument)invApp.ActiveDocument;
            Sheet oSheet = drawDoc.ActiveSheet;
            GeneralNotes oGenNotes = oSheet.DrawingNotes.GeneralNotes;

            // Check if keywords is empty then Get keywords from another GeneralNote
            if(valueKeywords == "")
            {
                try
                {
                    GeneralNote oGenNoteSelected = (GeneralNote)invApp.CommandManager.Pick(SelectionFilterEnum.kDrawingNoteFilter, "Select DrawingNote");
                    valueKeywords = GetTextFromNote(oGenNoteSelected.FormattedText, "Ký hiệu");
                    if(valueKeywords == "")
                    {
                        valueKeywords = "null";
                    }
                    else
                    {
                        propKeywords.Value = valueKeywords;
                        sFormattedText = "<StyleOverride FontSize = '0.28'><StyleOverride FontSize = '0.305' Underline = 'True'>" + "Chi tiết: " + valueTitle + "</StyleOverride>" +
                "\r\n" + "Ký hiệu: " + valueKeywords + "\r\n" + "Số lượng: " + valueQtyFinal + "\r\n" + "Vật liệu: " + valueComments + "</StyleOverride>";
                    }
                    oGenNotes.AddFitted(oGenNoteSelected.Position, sFormattedText);
                    oGenNoteSelected.Delete();
                }
                catch
                {

                }

                // Check GeneralNote exist
                bool CheckNGExist = false;
                int CountExist = 0;
                GeneralNote oGeneralNotTemp = null;
                foreach (GeneralNote oGeneralNote in drawDoc.ActiveSheet.DrawingNotes.GeneralNotes)
                {
                    if (oGeneralNote.FormattedText.ToUpper().IndexOf(valueTitle.ToUpper()) != 0)
                    {
                        oPosition = oGeneralNote.Position;
                        oGeneralNotTemp = oGeneralNote;
                        CountExist++;
                        CheckNGExist = true;
                    }
                }
                if (CheckNGExist == true) // Trường hợp đã tồn tại GeneralNote trùng khớp
                {
                    if(CountExist == 1) // Nếu có 1 trường hợp trùng khớp => Tạo mới, xoá cũ
                    {
                        oGenNotes.AddFitted(oPosition, sFormattedText);
                        oGeneralNotTemp.Delete();
                    }
                    else
                    {
                        GeneralNote oGenNoteSelected = (GeneralNote)invApp.CommandManager.Pick(SelectionFilterEnum.kDrawingNoteFilter, "Found " + CountExist + " results match. Pick one..");
                        oGenNotes.AddFitted(oGenNoteSelected.Position, sFormattedText);
                        oGenNoteSelected.Delete();
                    }
                }
                else
                {
                    double h, D;
                    h = oView.Height;
                    D = oView.Width;
                    TransientGeometry oTG = invApp.TransientGeometry;
                    oGenNotes.AddFitted(oTG.CreatePoint2d(oView.Position.X - D / 6, oView.Position.Y + h / 2 + 4), sFormattedText);
                }

            }
           
        }

        string GetTextFromNote(string oSource, string oFindText)
        {
            int oCheck, LenFindText, oStartPos, oEndPos;
            oCheck = oSource.IndexOf(oFindText);
            LenFindText = oFindText.Length;
            if(oCheck != 0)
            {
                oStartPos = oCheck - 1;
                oEndPos = oSource.IndexOf("<", oStartPos) - 1;
                oSource = Left(oSource, oEndPos);
                oSource = Right(oSource, oSource.Length - oStartPos);
                oSource = Right(oSource, oSource.Length - LenFindText - 2);
            }
            return oSource;
        }

        string Left(string input, int count)
        {
            return input.Substring(0, Math.Min(input.Length, count));
        }

        string Right(string input, int count)
        {
            return input.Substring(Math.Max(input.Length - count, 0), Math.Min(count, input.Length));
        }

        public void AddToUserInterface(string AddInClientID)
        {
            RibbonPanel panel;
            try
            {
                panel = invApp.UserInterfaceManager.Ribbons["Drawing"].RibbonTabs["id_TabPlaceViews"].RibbonPanels["id_UserToolTab"];
            }
            catch
            {
                panel = invApp.UserInterfaceManager.Ribbons["Drawing"].RibbonTabs["id_TabPlaceViews"].RibbonPanels.Add("Tool Panel", "id_UserToolTab", AddInClientID, "", false);
            }
            panel.CommandControls.AddButton(CreateDetailsButton);
        }

        public sealed class PictureDispConverter
        {
            [DllImport("OleAut32.dll", EntryPoint = "OleCreatePictureIndirect", ExactSpelling = true, PreserveSig = false)]
            private static extern stdole.IPictureDisp
                OleCreatePictureIndirect([MarshalAs(UnmanagedType.AsAny)] object picdesc, ref Guid iid, [MarshalAs(UnmanagedType.Bool)] bool fOwn);
            static Guid iPictureDispGuid = typeof(stdole.IPictureDisp).GUID;
            private static class PICTDESC
            {
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

            public static stdole.IPictureDisp ToIPictureDisp(
                System.Drawing.Icon icon)
            {
                PICTDESC.Icon pictIcon = new PICTDESC.Icon(icon);
                return OleCreatePictureIndirect(
                    pictIcon, ref iPictureDispGuid, true);
            }
            public static stdole.IPictureDisp ToIPictureDisp(
                System.Drawing.Bitmap bmp)
            {
                PICTDESC.Bitmap pictBmp = new PICTDESC.Bitmap(bmp);
                return OleCreatePictureIndirect(pictBmp, ref iPictureDispGuid, true);
            }
        }
    }

}
