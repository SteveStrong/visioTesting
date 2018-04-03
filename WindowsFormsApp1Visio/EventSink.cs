// EventSink.cs
// compile with: /doc:EventSink.xml
// <copyright>Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// <summary>This file contains the implementation for the EventSink 
// class which handles events from Visio.</summary>

using System;
using System.Diagnostics;
using Microsoft.Office.Interop.Visio;
using System.Diagnostics.CodeAnalysis;
using Visio = Microsoft.Office.Interop.Visio;

namespace WindowsFormsApp1Visio
{


    public enum DrawingVisioEvents : short
    {
        QueryCancelSelectionDelete = (short)Visio.VisEventCodes.visEvtCodeQueryCancelSelDel,
        QueryCancelMasterDelete = (short)Visio.VisEventCodes.visEvtCodeQueryCancelMasterDel,
        QueryCancelDocumentClose = (short)Visio.VisEventCodes.visEvtCodeQueryCancelDocClose,
        QueryCancelPageDelete = (short)Visio.VisEventCodes.visEvtCodeQueryCancelPageDel,
        QueryCancelQuit = (short)Visio.VisEventCodes.visEvtCodeQueryCancelQuit,

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



    //[SuppressMessage("Microsoft.Design", "CA1003:UseGenericEventHandlerInstances")]
    public delegate void VisioEventHandler(object sender, EventArgs e);

    /// <summary>The EventSink class handles events from Visio
    /// which are specified in the AddAdvise method.</summary>
    [System.Runtime.InteropServices.ComVisible(true)]
    public sealed class EventSink : IVisEventProc
    {

        private const string officePlanArgument = "/officeplan";
        private const string doubleClickCommand = "/cmd=1";

        /// <summary>Visio.Application object.</summary>
        private Visio.Application eventApplication;

        /// <summary>Visio.Document object.</summary>
        private Visio.Document eventDocument;

        /// <summary>Two FIFO queues are used to store added and deleted
        /// shape information while Visio events are being processed.</summary>
        private System.Collections.Queue shapeAddedQueue;
        private System.Collections.Queue shapeDeletedQueue;

        /// <summary>OnShapeAdd event is raised when a shape is
        ///  added to the drawing.</summary>
        public event VisioEventHandler OnShapeAdd;

        /// <summary>OnShapeDelete event is raised when a shape is
        ///  deleted from the drawing.</summary>
        public event VisioEventHandler OnShapeDelete;

        /// <summary>OnShapeDoubleClick event is raised when the user
        /// double-clicks a shape from the sample office furniture stencil.
        /// </summary>
        public event VisioEventHandler OnShapeDoubleClick;

        /// <summary>The EventSink default constructor creates the queues that
        /// will be used to hold added and deleted shapes for processing.</summary>
        public EventSink()
        {
            shapeAddedQueue = new System.Collections.Queue();
            shapeDeletedQueue = new System.Collections.Queue();
        }


        [CLSCompliant(false)]
        public void AddAdvise(Visio.Application callingApplication, Visio.Document callingDocument)
        {
            const string sink = "";
            const string targetArgs = "";

            // Save the application for setting the events.
            eventApplication = callingApplication;
            EventList applicationEvents = eventApplication.EventList;

            applicationEvents.AddAdvise((short)DrawingVisioEvents.AfterIdle, (IVisEventProc)this, sink, targetArgs);

            // Save the document for setting the events.
            eventDocument = callingDocument;
            EventList documentEvents = eventDocument.EventList;

            // Add events of interest.
            //setAddAdvise();

            documentEvents.AddAdvise((short)DrawingVisioEvents.AfterShapeAdded, (IVisEventProc)this, sink, targetArgs);
            //documentEvents.AddAdvise((unchecked((short)VisEventCodes.visEvtAdd) + (short)VisEventCodes.visEvtShape), (IVisEventProc)this, sink, targetArgs);

            return;
        }


        object IVisEventProc.VisEventProc(short eventCode, object source, int eventId, int eventSequenceNumber, object subject, object moreInfo)
        {

            Visio.IVApplication eventProcApplication = source as IVApplication;

            // Check for each event code that is handled.  The event
            // codes are a combination of an object and an action.
            // Only the events added in the SetAddAdvise method will
            // be sent to this method, and only those events need to
            // be included in this switch statement.
            Shape eventShape = null;
            if ((eventCode & (short)VisEventCodes.visEvtShape) > 0)
            {
                eventShape = (Shape)subject;
            }

            switch (eventCode)
            {

                case (short)DrawingVisioEvents.AfterShapeAdded:

                    // Handle the add-shape event.
                    handleShapeAdd(eventShape);
                    break;

                case (short)VisEventCodes.visEvtDel + (short)VisEventCodes.visEvtShape:

                    // Handle the delete-shape event.
                    handleShapeDelete(eventShape);
                    break;

                case (short)VisEventCodes.visEvtApp + (short)VisEventCodes.visEvtMarker:

                    // Handle this marker event.
                    handleMarker(eventProcApplication);
                    break;

                case (short)VisEventCodes.visEvtApp +(short)VisEventCodes.visEvtNonePending:

                    // Handle the no-events-pending event.
                    handleNonePending();
                    break;

                default:
                    break;
            }

            return null;
        }

        private void setAddAdvise()
        {

            // The Sink and TargetArgs values aren't needed.
            const string sink = "";
            const string targetArgs = "";

            EventList applicationEvents = eventApplication.EventList;
            EventList documentEvents = eventDocument.EventList;

            // Add the shape-added event to the document. The new shape
            // will be available for processing in the handler.  The
            // value for VisEventCodes.visEvtAdd cannot be
            // automatically converted to a short type, so the
            // unchecked function is used.  This allows the addition to
            // be done and returns a valid short value.
            documentEvents.AddAdvise(
                (unchecked((short)VisEventCodes.visEvtAdd) +
                (short)VisEventCodes.visEvtShape),
                (IVisEventProc)this, sink, targetArgs);

            // Add the before-shape-deleted event to the document.  This 
            // event will be raised when a shape is deleted from the
            // document. The deleted shape will still be available for
            // processing in the handler.
            documentEvents.AddAdvise(
                (short)VisEventCodes.visEvtDel +
                (short)VisEventCodes.visEvtShape,
                (IVisEventProc)this, sink, targetArgs);

            // Add marker events to the application.  This event
            // will be raised when a user double-clicks a shape from
            // the sample office furniture stencil.
            applicationEvents.AddAdvise(
                (short)VisEventCodes.visEvtApp +
                (short)VisEventCodes.visEvtMarker,
                (IVisEventProc)this, sink, targetArgs);

            // Add the none-pending event to the application.  This
            // event will be raised when Visio is idle.
            applicationEvents.AddAdvise(
                (short)VisEventCodes.visEvtApp +
                (short)VisEventCodes.visEvtNonePending,
                (IVisEventProc)this, sink, targetArgs);

            return;
        }

        /// <summary>The handleNonePending method is called when all Visio
        /// events have been processed.  The queued shape adds and deletes
        /// are processed during Visio's idle time.</summary>
        private void handleNonePending()
        {

            // Process the added-shapes queue.
            if (OnShapeAdd != null)
            {

                // Raise an OnShapeAdd event for each shape in the queue.
                while (shapeAddedQueue.Count > 0)
                {

                    OnShapeAdd(shapeAddedQueue.Dequeue(), new EventArgs());
                }
            }

            else
            {
                // There are no event listeners so just empty the queue.
                shapeAddedQueue.Clear();
            }

            // Process the deleted-shapes queue.
            if (OnShapeDelete != null)
            {

                // Raise an OnShapeDelete event for each shape in the 
                // queue.
                while (shapeDeletedQueue.Count > 0)
                {

                    OnShapeDelete(shapeDeletedQueue.Dequeue(),
                        new EventArgs());
                }
            }

            else
            {
                // There are no event listeners so just empty the queue.
                shapeDeletedQueue.Clear();
            }

            return;
        }

        private void handleShapeAdd(Shape addedShape)
        {
            // Add the shape to the queue.
            shapeAddedQueue.Enqueue(addedShape);
            return;
        }

        private void handleShapeDelete(IVShape deletedShape)
        {
            // Add the product ID to the queue.
            shapeDeletedQueue.Enqueue(deletedShape);
            return;
        }

        /// <summary>The handleMarker method is called when Visio raises
        /// a marker event. When the user double-clicks on a shape from the
        /// sample office furniture stencil the formula in the shape's 
        /// EventDblClick cell will run the QueueMarkerEvent addon which will
        /// raise a marker event.</summary>
        /// <param name="visioApplication">The Visio application that raised
        ///  this event.</param>
        //[SuppressMessage("Microsoft.Globalization", "CA1308:NormalizeStringsToUppercase")]
        private void handleMarker(IVApplication visioApplication)
        {

            string arguments;
            Shape targetShape;

            // If the arguments include /officeplan /cmd=1
            // then get a reference to the shape and raise the
            // OnShapeDoubleClick event

            if (OnShapeDoubleClick != null)
            {

                arguments = visioApplication.get_EventInfo((short)VisEventCodes.visEvtIdMostRecent);
                arguments = arguments.ToLowerInvariant();

                // If this marker event was caused by double-clicking a 
                // shape from the sample office furniture stencil then
                // raise an OnShapeDoubleClick event
                if ((arguments.IndexOf(officePlanArgument, StringComparison.Ordinal) >= 0) &&
                    (arguments.IndexOf(doubleClickCommand, StringComparison.Ordinal) >= 0))
                {

                    // Get a reference to this shape
                    targetShape = Utility.GetShapeFromArguments(visioApplication, arguments);

                    // Raise an OnShapeDoubleClick event for this shape.
                    OnShapeDoubleClick(targetShape, new EventArgs());
                }
            }

            return;
        }
    }
}
