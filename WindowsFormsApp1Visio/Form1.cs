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

    /// <summary>The EventSink class handles events from Visio
    /// which are specified in the AddAdvise method.</summary>
    [System.Runtime.InteropServices.ComVisible(true)]
    public sealed class VisioEventSink : Visio.IVisEventProc
    {


        /// <summary>Visio.Application object.</summary>
        private Visio.Application eventApplication;

        /// <summary>Visio.Document object.</summary>
        private Visio.Document eventDocument;

        [CLSCompliant(false)]
        public void AddAdvise(Visio.Application callingApplication, Visio.Document callingDocument)
        {
            const string sink = "";
            const string targetArgs = "";

            // Save the document for setting the events.
            eventDocument = callingDocument;
            Visio.EventList documentEvents = eventDocument.EventList;


            documentEvents.AddAdvise((short)DrawingEvents.AfterShapeAdded, (Visio.IVisEventProc)this, sink, targetArgs);

            //Visio.EventList ovEvents = m_ovTargetDoc.EventList;
            //CreateEvent(ovEvents, DrawingEvents.AfterShapeAdded);
        }




        object Visio.IVisEventProc.VisEventProc(short eventCode, object source, int eventId, int eventSequenceNumber, object subject, object moreInfo)
        {

            Visio.Shape eventShape = null;
            if ((eventCode & (short)Visio.VisEventCodes.visEvtShape) > 0)
            {
                eventShape = (Visio.Shape)subject;
            }

            switch (eventCode)
            {

                case (short)DrawingVisioEvents.AfterShapeAdded:

                    // Handle the add-shape event.
                    //handleShapeAdd(eventShape);
                    break;

                default:
                    break;
            }

            return null;
        }
    }


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
        private VisioEventSink visioEventSink;


        public Form1()
        {
            InitializeComponent();
            m_ovControl = axDrawingControl1;
            ultraTextEditor1.AcceptsReturn = true;
            m_ovControl.PageAdded += AxDrawingControl1_PageAdded;
        }


        private void onAddProductInformation(object sender, EventArgs e)
        {
            Report(e.ToString());
        }

        private void onRemoveProductInformation(object sender, EventArgs e)
        {
            Report(e.ToString());
        }

        private void onShapeDoubleClick(object sender, EventArgs e)
        {
            Report(e.ToString());
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

            visioEventSink = new VisioEventSink();
            visioEventSink.AddAdvise(m_ovTargetApp, m_ovTargetDoc);
            //setUpVisioDrawing();

            //m_ovControl.ShapeAdded += M_ovControl_ShapeAdded;

            //Visio.EventList ovEvents = m_ovTargetDoc.EventList;
            //CreateEvent(ovEvents, DrawingEvents.AfterShapeAdded);
            //EstablishEvent(ovEvents, DrawingEvents.BeforeShapeDeleted);
            //EstablishEvent(ovEvents, DrawingEvents.AfterParentChanged);
            return true;
        }

        private void M_ovControl_ShapeAdded(object sender, AxVisOcx.EVisOcx_ShapeAddedEvent e)
        {

            Report(e.shape.ToString());
        }

        object Visio.IVisEventProc.VisEventProc(short nEventCode, object pSourceObj, int nEventID, int nEventSeqNum, object pSubjectObj, object vMoreInfo)
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

        [CLSCompliant(false)]
        public Visio.Event CreateEvent(Visio.EventList ovEventList, DrawingEvents iEvent)
        {
            const string sink = "";
            const string targetArgs = "";
            try
            {
                return ovEventList.AddAdvise((short)iEvent, (Visio.IVisEventProc)this, sink, targetArgs);
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
