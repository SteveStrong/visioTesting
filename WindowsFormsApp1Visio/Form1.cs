using System;
using System.Windows.Forms;
using System.Collections;

using AxVisOcx = AxMicrosoft.Office.Interop.VisOcx;
using Visio = Microsoft.Office.Interop.Visio;
using System.Runtime.InteropServices;

namespace WindowsFormsApp1Visio
{
    public enum visTF : short { TRUE = -1, FALSE = 0 };
    public enum StencilOpenMode
    {
        ReadOnly,
        ReadWrite,
        ReadHidden,
        WriteHidden,
        FindToClose,
        FindToDelete
    }

    public enum DrawingEvents : short
    {
        QueryCancelSelectionDelete = (short)Visio.VisEventCodes.visEvtCodeQueryCancelSelDel,
        QueryCancelMasterDelete = (short)Visio.VisEventCodes.visEvtCodeQueryCancelMasterDel,
        QueryCancelDocumentClose = (short)Visio.VisEventCodes.visEvtCodeQueryCancelDocClose,
        QueryCancelPageDelete = (short)Visio.VisEventCodes.visEvtCodeQueryCancelPageDel,
        QueryCancelQuit = (short)Visio.VisEventCodes.visEvtCodeQueryCancelQuit,

        //		AfterMouseDown = (short)Visio.VisEventCodes.visEvtCodeMouseDown,
        //		AfterMouseMove = (short)Visio.VisEventCodes.visEvtCodeMouseMove,
        //		AfterMouseUp = (short)Visio.VisEventCodes.visEvtCodeMouseUp,

        AfterKeyPress = (short)Visio.VisEventCodes.visEvtCodeKeyPress,
        AfterKeyDown = (short)Visio.VisEventCodes.visEvtCodeKeyDown,
        AfterKeyUp = (short)Visio.VisEventCodes.visEvtCodeKeyUp,


        AfterSelectionChanged = (short)Visio.VisEventCodes.visEvtCodeWinSelChange,
        AfterMarker = (unchecked((short)Visio.VisEventCodes.visEvtApp) + (short)Visio.VisEventCodes.visEvtMarker),
        AfterWindowActivate = (unchecked((short)Visio.VisEventCodes.visEvtApp) + (short)Visio.VisEventCodes.visEvtWinActivate),
        AfterIdle = (unchecked((short)Visio.VisEventCodes.visEvtApp) + (short)Visio.VisEventCodes.visEvtNonePending),
        BeforeTextEdit = (short)Visio.VisEventCodes.visEvtCodeShapeBeforeTextEdit,
        AfterTextEdit = (short)Visio.VisEventCodes.visEvtCodeShapeExitTextEdit,
        AfterTextChanged = (unchecked((short)Visio.VisEventCodes.visEvtMod) + (short)Visio.VisEventCodes.visEvtText),
        AfterCellChanged = (unchecked((short)Visio.VisEventCodes.visEvtMod) + (short)Visio.VisEventCodes.visEvtCell),
        AfterParentChanged = (short)Visio.VisEventCodes.visEvtCodeShapeParentChange,
        BeforePageTurn = (short)Visio.VisEventCodes.visEvtCodeBefWinPageTurn,
        AfterPageTurn = (short)Visio.VisEventCodes.visEvtCodeWinPageTurn,
        AfterPageChanged = (unchecked((short)Visio.VisEventCodes.visEvtMod) + (short)Visio.VisEventCodes.visEvtPage),
        BeforeApplicationQuit = (short)Visio.VisEventCodes.visEvtBeforeQuit,
        AfterDocumentOpened = (unchecked((short)Visio.VisEventCodes.visEvtAdd) + (short)Visio.VisEventCodes.visEvtDoc),
        BeforeDocumentClosed = (unchecked((short)Visio.VisEventCodes.visEvtDel) + (short)Visio.VisEventCodes.visEvtDoc),
        AfterPageAdded = (unchecked((short)Visio.VisEventCodes.visEvtAdd) + (short)Visio.VisEventCodes.visEvtPage),
        BeforePageDeleted = (unchecked((short)Visio.VisEventCodes.visEvtDel) + (short)Visio.VisEventCodes.visEvtPage),
        AfterShapeAdded = (unchecked((short)Visio.VisEventCodes.visEvtAdd) + Visio.VisEventCodes.visEvtShape),
        BeforeShapeDeleted = (unchecked((short)Visio.VisEventCodes.visEvtDel) + (short)Visio.VisEventCodes.visEvtShape),
        AfterConnectionAdded = (unchecked((short)Visio.VisEventCodes.visEvtAdd) + (short)Visio.VisEventCodes.visEvtConnect),
        BeforeConnectionDeleted = (unchecked((short)Visio.VisEventCodes.visEvtDel) + (short)Visio.VisEventCodes.visEvtConnect),
        BeforeSelectionDeleted = (short)Visio.VisEventCodes.visEvtCodeBefSelDel,
        BeforeWindowSelectionDeleted = (short)Visio.VisEventCodes.visEvtCodeBefWinSelDel,
    }

    //[System.Runtime.InteropServices.ComVisible(true)]
    [CLSCompliant(false)]
    public partial class Form1 : Form, Visio.IVisEventProc
    {
        protected Visio.Application m_ovTargetApp = null;
        protected Visio.Document m_ovTargetDoc = null;
        protected AxVisOcx.AxDrawingControl m_ovControl = null;
        protected Hashtable m_ovEvents = new Hashtable();


        private Visio.Application visioApplication;
        private Visio.Document visioDocument;
        private int alertResponse = 0;
        private string applicationName = "";
        private EventSink visioEventSink;

        public Form1()
        {
            InitializeComponent();
            m_ovControl = axDrawingControl1;
            ultraTextEditor1.AcceptsReturn = true;
            m_ovControl.PageAdded += AxDrawingControl1_PageAdded;
            //m_ovControl.VisibleChanged += M_ovControl_VisibleChanged;
        }

        private void M_ovControl_VisibleChanged(object sender, EventArgs e)
        {
            setUpVisioDrawing();
        }

        private void setUpVisioDrawing()
        {

            double pageLeft;
            double pageTop;
            double pageWidth;
            double pageHeight;

            try
            {

                // Cache the application name for use in message box captions.
                applicationName = m_ovControl.Window.Application.Name;

                // Cache the AlertResponse setting in a member variable.
                alertResponse = m_ovControl.Window.Application.AlertResponse;

                // Hide all built-in docked windows (shape search, 
                // custom properties (shape data), etc.).
                //for (int i = m_ovControl.Window.Windows.Count; i > 0; i--)
                //{
                //    Visio.Window visWindow;
                //    int windowType;

                //    visWindow = m_ovControl.Window.Windows.get_ItemEx(i);
                //    windowType = visWindow.Type;

                //    if (windowType == (int)Visio.VisWinTypes.visAnchorBarBuiltIn)
                //    {

                //        switch (visWindow.ID)
                //        {
                //            case (int)Visio.VisWinTypes.visWinIDCustProp:
                //            case (int)Visio.VisWinTypes.visWinIDDrawingExplorer:
                //            case (int)Visio.VisWinTypes.visWinIDMasterExplorer:
                //            case (int)Visio.VisWinTypes.visWinIDPanZoom:
                //            case (int)Visio.VisWinTypes.visWinIDShapeSearch:
                //            case (int)Visio.VisWinTypes.visWinIDSizePos:
                //            case (int)Visio.VisWinTypes.visWinIDStencilExplorer:

                //                visWindow.Visible = false;
                //                break;

                //            default:
                //                break;
                //        }
                //    }
                //}

                // Use the Visio window to set the visible user
                // interface parts of the window.
                Visio.Window targetWindow;
                targetWindow = (Visio.Window)m_ovControl.Window;

                //targetWindow.ShowRulers = 0;
                //targetWindow.ShowPageTabs = false;
                //targetWindow.ShowScrollBars = 0;
                //targetWindow.ShowGrid = 0;
                //targetWindow.Zoom = 1.00;

                // Position the furniture shapes relative to the page.
                targetWindow.GetViewRect(out pageLeft, out pageTop, out pageWidth, out pageHeight);



                // Start the event sink.
                initializeEventSink();
            }

            catch (COMException error)
            {

                // Display the error.
                Utility.DisplayException(Strings.ComErrorMessage,error, alertResponse);
                throw;
            }

            return;
        }

        private void initializeEventSink()
        {

            try
            {

                // Release the previous event sink.
                visioEventSink = null;

                // Create an event sink to hook up events to the Visio
                // application and document.
                visioEventSink = new EventSink();
                visioApplication = (Visio.Application)m_ovControl.Window.Application;
                visioDocument = (Visio.Document)m_ovControl.Document;

                visioEventSink.AddAdvise(visioApplication, visioDocument);

                // Listen to shape add events from the Visio event sink.
                // OnAddProductInformation will be called when a shape is added.
                visioEventSink.OnShapeAdd += new VisioEventHandler(onAddProductInformation);

                // Listen to shape delete events from the Visio event sink.
                // OnRemoveProductInformation will be called when a shape is deleted.
                visioEventSink.OnShapeDelete += new VisioEventHandler(onRemoveProductInformation);

                // Listen to marker events raised when the user double-clicks a shape.
                // OnShapeDoubleClick will be called when the user double-clicks
                // a shape whose double-click event runs the QueueMarkerEvent addon.
                visioEventSink.OnShapeDoubleClick += new VisioEventHandler(onShapeDoubleClick);
            }

            catch (COMException error)
            {

                // Display the error.
                Utility.DisplayException(Strings.ComErrorMessage, error, alertResponse);
                throw;
            }
        }

        private void onAddProductInformation(object sender, EventArgs e)
        {
        }

        private void onRemoveProductInformation(object sender, EventArgs e)
        {
        }

        private void onShapeDoubleClick(object sender, EventArgs e)
        {
        }

        void Report(string text)
        {
            this.ultraTextEditor1.AppendText(text);
            this.ultraTextEditor1.AppendText("\n");
        }

        void ReportException(Exception e)
        {
            Report(e.Message);
        }

        private bool StartEvents()
        {
            LoadDocument(@"C:\Users\Steve\Documents\visioTesting\WindowsFormsApp1Visio\sample.vsd");
            if (m_ovTargetDoc == null)
                return false;

            setUpVisioDrawing();

            //m_ovControl.ShapeAdded += M_ovControl_ShapeAdded;

            //Visio.EventList ovEvents = m_ovTargetDoc.EventList;
            //EstablishEvent(ovEvents, DrawingEvents.AfterShapeAdded);
            //EstablishEvent(ovEvents, DrawingEvents.BeforeShapeDeleted);
            //EstablishEvent(ovEvents, DrawingEvents.AfterParentChanged);
            return true;
        }

        private void M_ovControl_ShapeAdded(object sender, AxVisOcx.EVisOcx_ShapeAddedEvent e)
        {

            Report(e.shape.ToString());
        }

        public dynamic VisEventProc(short nEventCode, object pSourceObj, int nEventID, int nEventSeqNum, object pSubjectObj, object vMoreInfo)
        {
            //Debug.WriteLine( string.Format("Event Code: 0x{0:X}", nEventCode));

            //very critical to prevent other applications from reacting
            //to events ment for this one.
            Visio.Application oApp = pSourceObj as Visio.Application;


            try
            {
                switch (nEventCode)
                {
                    case (short)DrawingEvents.BeforeApplicationQuit:
                        break;
                    case (short)DrawingEvents.AfterSelectionChanged:
                        Visio.Window ovWin = pSubjectObj as Visio.Window;
                        break;
                    case (short)DrawingEvents.BeforeTextEdit:
                        break;
                    case (short)DrawingEvents.AfterTextEdit:
                        break;
                    case (short)DrawingEvents.AfterTextChanged:
                        break;
                    case (short)DrawingEvents.BeforePageTurn:
                        break;
                    case (short)DrawingEvents.AfterPageTurn:
                        break;
                    case (short)DrawingEvents.AfterPageChanged:
                        break;
                    case (short)DrawingEvents.AfterParentChanged:
                        break;
                    case (short)DrawingEvents.AfterDocumentOpened:
                        break;
                    case (short)DrawingEvents.BeforeDocumentClosed:
                        break;
                    case (short)DrawingEvents.AfterPageAdded:
                        break;
                    case (short)DrawingEvents.BeforePageDeleted:
                        break;
                    case (short)DrawingEvents.AfterShapeAdded:
                        var m_iLastScope = m_ovTargetApp.CurrentScope;
                        switch (m_iLastScope)
                        {
                            case 1166: // Duplicate from automation call shape
                                break; // Shape should not be added
                            case 1184: // control drag shapes
                            case 1024: // Duplicate shape
                            case 1017: // undo shape
                                break;
                            case 1018: // redo shape
                            case 1022: // Paste shape
                                break;
                            case 0: // undo shape sort of
                                break;
                            default:
                                break;
                        }
                        break;
                    case (short)DrawingEvents.BeforeWindowSelectionDeleted:
                        break;
                    case (short)DrawingEvents.BeforeSelectionDeleted:
                        m_iLastScope = m_ovTargetApp.CurrentScope;
                        switch (m_iLastScope)
                        {
                            case 1023: // Delete key
                                break;
                            case 1020: // Cut shape
                            case 1017: // undo shape
                            case 1018: // redo shape
                                break;
                            case 1486: // deleted from page on moved
                            default:
                                break;
                        }
                        break;
                    case (short)DrawingEvents.BeforeShapeDeleted:
                        m_iLastScope = m_ovTargetApp.CurrentScope;
                        break;
                    case (short)DrawingEvents.AfterConnectionAdded:
                        break;
                    case (short)DrawingEvents.BeforeConnectionDeleted:
                        break;
                    case (short)DrawingEvents.AfterKeyPress:
                        break;
                    case (short)DrawingEvents.AfterKeyDown:
                        break;
                    case (short)DrawingEvents.AfterKeyUp:
                        break;

                    case (short)DrawingEvents.QueryCancelSelectionDelete:
                        break;
                    case (short)DrawingEvents.QueryCancelMasterDelete:
                        break;
                    case (short)DrawingEvents.QueryCancelPageDelete:
                        break;
                    case (short)DrawingEvents.QueryCancelDocumentClose:
                        break;
                    case (short)DrawingEvents.QueryCancelQuit:
                        break;

                    default:
                        return null;
                }
            }
            catch (Exception e)
            {
                ReportException(e);
            }
            return null;
        }



        private void AxDrawingControl1_PageAdded(object sender, AxVisOcx.EVisOcx_PageAddedEvent e)
        {
            Report(e.page.NameU);
        }


        public void EstablishEvent(Visio.EventList ovEventList, DrawingEvents iEvent, bool bProcess = true)
        {
            if (bProcess)
                EnableEvent(ovEventList, iEvent);
            else
                DisableEvent(ovEventList, iEvent, true);
        }

        public Visio.Event EnableEvent(Visio.EventList ovEventList, DrawingEvents iEvent)
        {
            Visio.Event ovEvent = null;
            string sKey = iEvent.ToString();

            if (m_ovEvents.ContainsKey(sKey) == true)
                ovEvent = m_ovEvents[sKey] as Visio.Event;
            else
            {
                ovEvent = CreateEvent(ovEventList, iEvent);
                if (ovEvent == null)
                    return null;

                m_ovEvents.Add(sKey, ovEvent);
            }
            ovEvent.Enabled = (short)visTF.TRUE;
            return ovEvent;
        }

        public Visio.Event DisableEvent(Visio.EventList ovEventList, DrawingEvents iEvent, bool bRemove)
        {
            Visio.Event ovEvent = null;
            string sKey = iEvent.ToString();

            if (m_ovEvents.ContainsKey(sKey) == true)
            {
                ovEvent = m_ovEvents[sKey] as Visio.Event;
                if (bRemove)
                    m_ovEvents.Remove(sKey);

                try
                {
                    ovEvent.Enabled = (short)visTF.FALSE;
                }
                catch { }
            }
            return ovEvent;
        }


        public Visio.Event CreateEvent(Visio.EventList ovEventList, DrawingEvents iEvent)
        {
            try
            {
                return ovEventList.AddAdvise((short)iEvent, this, "", "");
            }
            catch (Exception e)
            {
                ReportException(e);
            }
            return null;
        }


        private Visio.Document LoadDocument(string sPath)
        {

            Name = sPath;

            try
            {
                m_ovControl.Src = sPath;
            }
            catch (Exception e)
            {
                ReportException(e);
                return null;
            }

            m_ovTargetDoc = m_ovControl.Document;
            m_ovTargetApp = m_ovTargetDoc.Application;

            m_ovTargetApp.Settings.ZoomOnRoll = true;
            m_ovTargetApp.Settings.CenterSelectionOnZoom = true;
            m_ovTargetApp.Settings.ConnectorSplittingEnabled = true;
            m_ovTargetApp.Settings.DrawingAids = true;
            m_ovTargetApp.Settings.HigherQualityShapeDisplay = true;
            m_ovTargetApp.Settings.SmoothDrawing = true;
            m_ovTargetApp.EventsEnabled = (short)visTF.TRUE;

            return m_ovTargetDoc;
        }


        private void ultraButton1_Click(object sender, EventArgs e)
        {
            StartEvents();
        }


    }
}
