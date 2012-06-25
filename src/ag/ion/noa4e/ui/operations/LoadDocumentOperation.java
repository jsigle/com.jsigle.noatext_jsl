/****************************************************************************
 *                                                                          *
 * NOAText_jsl based upon NOA (Nice Office Access) / noa-libre              *
 * ------------------------------------------------------------------------ *
 *                                                                          *
 * The Contents of this file are made available subject to                  *
 * the terms of GNU General Public License Version 2.1                      *
 *                                                                          * 
 * GNU General Public License Version 2.1                                   *
 * ======================================================================== *
 * Portions Copyright 2012 by Joerg Sigle                                   *
 * Copyright 2003-2006 by IOn AG                                            *
 *                                                                          *
 * This program is free software: you can redistribute it and/or modify     *
 * it under the terms of the GNU General Public License as published by     *
 * the Free Software Foundation, either version 2.1 of the License.         *
 *                                                                          *
 *  This program is distributed in the hope that it will be useful,         *
 *  but WITHOUT ANY WARRANTY; without even the implied warranty of          *
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the           *
 *  GNU General Public License for more details.                            *
 *                                                                          *
 *  You should have received a copy of the GNU General Public License       *
 *  along with this program.  If not, see <http://www.gnu.org/licenses/>.   *
 *                                                                          *
 * Contact us:                                                              *
 *  http://www.jsigle.com                                                   *
 *  http://www.ql-recorder.com                                              *
 *  http://code.google.com/p/noa-libre                                      *
 *  http://www.ion.ag                                                       *
 *  http://ubion.ion.ag                                                     *
 *  info@ion.ag                                                             *
 *                                                                          *
 * Please note: Previously, versions of the NOA library provided by         *
 * www.ion.ag and the noa-libre project carried a licensing remark          *
 * that made them available under the LGPL. However, they include portions  *
 * obtained from the YaBS project, licensed under GPL. Consequently, NOA    *
 * should have been licensed under the GPL, not LGPL, given that no special *
 * permission of the authors of YaBS for LGPL licensing had been obtained.  *
 * To point out the possible problem, I'm providing the files where I added *
 * contributions under the GPL for now. This move is always allowed for     *
 * LPGL licensed material. 20120623js                                       * 
 *                                                                          *
 ****************************************************************************/

/****************************************************************************
 * To Do:
 * See experiment / comments by Joerg Sigle below.
 * Review what's happening inside called code (i.e. in OpenOffice/LibreOffice
 * libraries); why does it all work when the office window is not generated
 * inside a frame, but on its own? Apply necessary corrections so that the
 * same grade of robustness is achieved in a frame.
 * Possibly, this version of the library should get a new revision number, currently used is 11672.
 ****************************************************************************/

/*
 * Last changes made by $Author: jsigle $, $Date: 2012-06-23 14:38:00 +0100 (Su, 23 Jun 2012) $
 */
package ag.ion.noa4e.ui.operations;

import ag.ion.bion.officelayer.application.IOfficeApplication;

import ag.ion.bion.officelayer.desktop.GlobalCommands; //20120222js: To support disabling menus items and open dialogs which would cause the Elexis/OO to hang
import ag.ion.bion.officelayer.desktop.IFrame;

import ag.ion.bion.officelayer.document.IDocument;
import ag.ion.bion.officelayer.document.IDocumentDescriptor;
import ag.ion.bion.officelayer.document.IDocumentService; //20120222js: To support disabling menus items and open dialogs which would cause the Elexis/OO to hang
import ag.ion.bion.officelayer.text.ITextDocument;
import ag.ion.noa.frame.ILayoutManager; //20120222js: To support disabling menus items and open dialogs which would cause the Elexis/OO to hang

import com.sun.star.frame.XLayoutManager; //20120222js: To support disabling menus items and open dialogs which would cause the Elexis/OO to hang
import com.sun.star.frame.XFrame; //20120222js: To support disabling menus items and open dialogs which would cause the Elexis/OO to hang
import com.sun.star.uno.UnoRuntime; //20120222js: To support disabling menus items and open dialogs which would cause the Elexis/OO to hang
import com.sun.star.beans.XPropertySet; //20120222js: To support disabling menus items and open dialogs which would cause the Elexis/OO to hang
import com.sun.star.ui.UIElementType; //20120222js: To support disabling menus items and open dialogs which would cause the Elexis/OO to hang
import com.sun.star.frame.FrameSearchFlag; //20120222js: To support disabling menus items and open dialogs which would cause the Elexis/OO to hang

import org.eclipse.core.runtime.IProgressMonitor;

import org.eclipse.jface.operation.IRunnableWithProgress;

import java.io.InputStream;

import java.lang.reflect.InvocationTargetException;

import java.net.URL;

/**
 * Operation in order to load a OpenOffice.org document.
 * 
 * @author Andreas Br�ker
 * @version $Revision: 11672 $
 */
public class LoadDocumentOperation implements IRunnableWithProgress {

	private IOfficeApplication officeApplication = null;
	private IFrame frame = null;
	private IDocumentDescriptor documentDescriptor = null;

	private String documentType = null;

	private URL url = null;

	private InputStream inputStream = null;

	private boolean isSubTask = false;
	private boolean updateProgressMonitor = true;
	private boolean useStream = false;

	private InternalThread internalThread = null;

	// ----------------------------------------------------------------------------
	/**
	 * Internal thread class in order load a a OpenOffice.org text document.
	 * 
	 * @author Andreas Br�ker
	 */
	private class InternalThread extends Thread {

		private Exception exception = null;

		private IDocument document = null;

		private boolean done = false;

		// ----------------------------------------------------------------------------
		/**
		 * Executes thread logic.
		 * 
		 * Joerg Sigle added code for progress monitoring, to examine the stability problems observed in Elexis,
		 * and to improve stability by removing less desired menu options, by closing dialogs, and by docking undocked toolbars,
		 * as well as comprehensive comments on available and tested commands and results etc.
		 * 
		 * @author Joerg Sigle
		 * @date 22.02.2012
		 * 
		 * @author Andreas Br�ker
		 */
		public void run() {
			System.out.println("LoadDocumentOperation: InternalThread: run: begins");
			try {
				System.out.println("LoadDocumentOperation: InternalThread: run: Trying to do work...");
				if (useStream || inputStream != null) {
					System.out.println("LoadDocumentOperation: InternalThread: run: useStream || inputStream != null");
					InputStream inputStream = null;
					if (LoadDocumentOperation.this.inputStream != null)
						inputStream = LoadDocumentOperation.this.inputStream;
					else
						inputStream = url.openStream();
					if (frame != null)
						document = officeApplication.getDocumentService().loadDocument(frame, inputStream, documentDescriptor);
					else
						document = officeApplication.getDocumentService().loadDocument(inputStream, documentDescriptor);
					try {
						inputStream.close();
					} catch (Throwable throwable) {
						// do not consume
					}
				} else {
					System.out.println("LoadDocumentOperation: InternalThread: run: NOT (useStream || inputStream != null)");
					System.out.println("LoadDocumentOperation: InternalThread: run: Status before loadDocument()...");

					if (officeApplication == null)
						System.out.println("LoadDocumentOperation: InternalThread: run: officeApplication==null");
					else
						System.out.println("LoadDocumentOperation: InternalThread: run: officeApplication=" + officeApplication.toString());
					if (url == null)
						System.out.println("LoadDocumentOperation: InternalThread: run: url==null");
					else
						System.out.println("LoadDocumentOperation: InternalThread: run: url=" + url.toString());
					if (frame == null)
						System.out.println("LoadDocumentOperation: InternalThread: run: frame==null");
					else
						System.out.println("LoadDocumentOperation: InternalThread: run: frame=" + frame.toString());
					if (documentDescriptor == null)
						System.out.println("LoadDocumentOperation: InternalThread: run: documentDescriptor==null");
					else
						System.out.println("LoadDocumentOperation: InternalThread: run: documentDescriptor=" + documentDescriptor.toString());

					// 201202222250js:
					// Wenn ich immer NICHT nach frame lade, dann steht Elexis
					// nach dem Beginn des Ladens, wenn in OO ein Fenster wie
					// Druckformatvorlage offen war,
					// zwar auch - aber nach dem Beenden von Elexis (mit Fehler:
					// Elexis reagiert nicht) folgt dann kein identischer
					// Abbruch von OO,
					// sondern OO ist mit dem Dokument als eigenes Fenster
					// vorhanden, und das offene Druckformatvorlage-Fenster
					// auch.
					// Also vermutlich will das Druckformatvorlagen-Fenster vor
					// allem nicht in den Frame?
					// Oder, das nicht binden an den Frame schützt das
					// halboffene OpenOffice vor den Folgen des Abschusses des
					// stehenden Elexis?
					System.out.println("LoadDocumentOperation: InternalThread: run: officeApplication.getDesktopService()="
							+ officeApplication.getDesktopService());
					System.out.println("LoadDocumentOperation: InternalThread: run: officeApplication.getDesktopService().getFramesCount()="
							+ officeApplication.getDesktopService().getFramesCount());

					// Es gibt unter getApplicationInfo() auch Methoden, um
					// einzelne Infos oder z.B. die Info, ob der initial config
					// wizard gelaufen ist etc. herauszulesen:
					// System.out.println("LoadDocumentOperation: InternalThread: run: officeApplication.getApplicationInfo().dumpInfo()=");
					// officeApplication.getApplicationInfo().dumpInfo();

					if (frame != null) {

						System.out
								.println("LoadDocumentOperation: InternalThread: run: TO DO: !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
						System.out
								.println("LoadDocumentOperation: InternalThread: run: TO DO: MOVE THE disableDispatch and layoutManager stuff to a better place!");
						System.out
								.println("LoadDocumentOperation: InternalThread: run: TO DO: (Or maybe not - LayoutManager stuff only effective after the docLoad?)");
						System.out
								.println("LoadDocumentOperation: InternalThread: run: TO DO: (Maybe due to some \"Load settings with document...\" option, though.)");
						System.out
								.println("LoadDocumentOperation: InternalThread: run: TO DO: !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");

						// This could be used if we wouldn't know the frame
						// already:
						// IFrame frame = document.getFrame();

						// The following tweaks alltogether do NOT suffice to
						// avoid or close all floating dialog windows,
						// which might be (re-)open(ed) when OpenOffice starts
						// inside Elexis, and thereby halt the
						// OpenOffice within Elexis startup for some reason or
						// another (probably because the main OO window
						// cannot be localized as expected, or its services are
						// unavailable??/unaccessible?, while another dialog is
						// there).
						// Testcase: Select Patient. Add new letter for him
						// using any template. Go to top left position, insert
						// table.
						// Put text cursor into that table - will open table
						// dialog. Close "Briefe" view (thereby saving the
						// letter to
						// the database in exactly this state).
						// Reopen the letter - this will redisplay the letter,
						// the table, with the cursor inside, and thus also
						// create the table dialog immediately. No matter
						// whether it couldn't be closed by the following
						// commands,
						// or not - it hinders the continued startup. If in the
						// same elexis as before, it might even still be seen
						// blinking up before the whole Elexis/OpenOffice frame
						// stalls. 201202232313js

						// 20120222js: Added this in, adopted from the examples:
						// http://www.usegroup.de/software/noa/index.php/Hello_World_in_SWT/JFace
						// http://ubion.ion.ag/mainForumFolder/oiep_forum/0053/
						// Now it is time to disable two commands in the frame
						// This comes from snippet14
						frame.disableDispatch(GlobalCommands.NEW_DOCUMENT);
						// frame.disableDispatch(GlobalCommands.NEW_MENU); ??js
						frame.disableDispatch(GlobalCommands.CLOSE_DOCUMENT);
						frame.disableDispatch(GlobalCommands.QUIT_APPLICATION);
						// frame.disableDispatch(GlobalCommands.PRINT_DOCUMENT);
						// frame.disableDispatch(GlobalCommands.PRINT_DOCUMENT_DIRECT);

						// The following commands are not bad as such - BUT, if
						// the dialogs they open remain open throughout the
						// closing of OpenOffice,
						// they will remain open, be opened again when it is
						// next started, and thereby cause the Elexis/OO bridge
						// to hang when opening a doc.
						// (Just like almost any OpenOffice dialog open on
						// startup time)
						// So I inhibit the dialog opening commands for dialogs
						// which will probably NOT be used within OpenOffice.
						frame.disableDispatch(GlobalCommands.OPEN_HYPERLINK_DIALOG);
						frame.disableDispatch(GlobalCommands.EDIT_HYPERLINK);
						frame.disableDispatch(GlobalCommands.OPEN_DRAW_TOOLBAR);
						frame.disableDispatch(GlobalCommands.OPEN_NAVIGATOR);
						frame.disableDispatch(GlobalCommands.OPEN_GALLERY);
						frame.disableDispatch(GlobalCommands.OPEN_DATASOURCES);
						frame.disableDispatch(GlobalCommands.OPEN_STYLE_SHEET); // Schützt
																				// NICHT
																				// vor
																				// bereits
																				// extern
																				// geöffnetem
																				// F11
																				// Formatvorlagendialog
						// frame.disableDispatch(GlobalCommands.OPEN_HELP);
						frame.disableDispatch(GlobalCommands.OPEN_ASSISTS);
						// frame.disableDispatch(GlobalCommands.OPEN_VERSION_DIALOG);
						// //Diese hier stören vielleicht das Wiederöffnen,
						// frame.disableDispatch(GlobalCommands.OPEN_OTHER_FIELDS_DIALOG);
						// //enthalten aber an sich sinnvolle Funktionalitäten,
						// frame.disableDispatch(GlobalCommands.OPEN_INSERT_SCRIPT_DIALOG);
						// //bzw. solche, die hoffentlich nur Leute verwenden,
						// frame.disableDispatch(GlobalCommands.OPEN_FONT_DIALOG);
						// //welche sich im Fall eines Stehenbleibens zu helfen
						// wissen.
						// Es gibt noch einige weitere in GlobalCommands; und
						// vermutlich noch weitere via uno (siehe, woher
						// GlobalCommands sich speist).
						// Die lasse ich jetzt aber mal unbesehen stehen.

						frame.updateDispatches(); // 201202230410js: Only
													// thereafter, the
													// disableDispatch() has
													// some effect!

						// Das folgende wirkt hier oben anscheinend nicht bis
						// nach dem loadDocument():
						ILayoutManager layoutManager = frame.getLayoutManager();
						layoutManager.hideAll(); // 201202230412js: This becomes
													// effective immediately.
													// But maybe it doesn't help
													// before the load document!
													// Maybe it only affects the
													// doc that will immed be
													// unloaded.

						/*
						 * layoutManager.showElement(ILayoutManager.);
						 * layoutManager.showElement(ILayoutManager.);
						 * layoutManager.showElement(ILayoutManager.); public
						 * static String[] ALL_BARS_URLS = new String[] {
						 * URL_MENUBAR, URL_STATUSBAR, URL_TOOLBAR,
						 * URL_TOOLBAR_ALIGNMENTBAR, URL_TOOLBAR_ARROWSHAPES,
						 * URL_TOOLBAR_BASICSHAPES, URL_TOOLBAR_CALLOUTSHAPES,
						 * URL_TOOLBAR_DRAWBAR, URL_TOOLBAR_DRAWOBJECTBAR,
						 * URL_TOOLBAR_EXTRUSIONOBJECTBAR,
						 * URL_TOOLBAR_FONTWORKOBJECTBAR,
						 * URL_TOOLBAR_FONTWORKSHAPETYPES,
						 * URL_TOOLBAR_FORMATOBJECTBAR,
						 * URL_TOOLBAR_FORMCONTROLS, URL_TOOLBAR_FORMDESIGN,
						 * URL_TOOLBAR_FORMFILTERBAR,
						 * URL_TOOLBAR_FORMSNAVIGATIONBAR,
						 * URL_TOOLBAR_FORMSOBJECTBAR,
						 * URL_TOOLBAR_FORMSTEXTOBJECTBAR,
						 * URL_TOOLBAR_FULLSCREENBAR,
						 * URL_TOOLBAR_GRAPHICOBJECTBAR, URL_TOOLBAR_INSERTBAR,
						 * URL_TOOLBAR_INSERTCELLSBAR,
						 * URL_TOOLBAR_INSERTOBJECTSBAR,
						 * URL_TOOLBAR_MEDIAOBJECTSBAR,
						 * URL_TOOLBAR_MOREFORMCONTROLS, URL_TOOLBAR_PREVIEWBAR,
						 * URL_TOOLBAR_STANDARDBAR, URL_TOOLBAR_STARSHAPES,
						 * URL_TOOLBAR_SYMBOLSHAPES, URL_TOOLBAR_TEXTOBJECTBAR,
						 * URL_TOOLBAR_VIEWERBAR };
						 */

						// System.out.println("LoadDocumentOperation: InternalThread: run: WARNING: !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
						// System.out.println("LoadDocumentOperation: InternalThread: run: WARNING: constructNewDocument instead of loadDocument used for debugging purposes!");
						// System.out.println("LoadDocumentOperation: InternalThread: run: WARNING: !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
						// document =
						// officeApplication.getDocumentService().constructNewDocument(frame,
						// IDocument.WRITER, documentDescriptor);

						XFrame xFrame = frame.getXFrame();
						XPropertySet xPropSet = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xFrame);
						XLayoutManager xLayoutManager = (XLayoutManager) UnoRuntime.queryInterface(XLayoutManager.class,
								xPropSet.getPropertyValue("LayoutManager"));
						xLayoutManager.dockAllWindows(com.sun.star.ui.UIElementType.FLOATINGWINDOW);
						xLayoutManager.dockAllWindows(com.sun.star.ui.UIElementType.TOOLBAR);
						xLayoutManager.dockAllWindows(com.sun.star.ui.UIElementType.MENUBAR);
						xLayoutManager.dockAllWindows(com.sun.star.ui.UIElementType.POPUPMENU);
						xLayoutManager.dockAllWindows(com.sun.star.ui.UIElementType.STATUSBAR);
						xLayoutManager.dockAllWindows(com.sun.star.ui.UIElementType.PROGRESSBAR);
						xLayoutManager.dockAllWindows(com.sun.star.ui.UIElementType.UNKNOWN);
						// xLayoutManager.dockAllWindows(com.sun.star.ui.UIElementType.TOOLPANEL);

						// System.out.println("LoadDocumentOperation: xFrame.getName()"+xFrame.getName());

						// System.out.println("findFrame(\"Formatvorlagen\"):"+xFrame.findFrame("Formatvorlagen",com.sun.star.frame.FrameSearchFlag.ALL).getName());

						System.out.println("LoadDocumentOperation: PRE");
						// document =
						// officeApplication.getDocumentService().loadDocument(frame,
						// url.toString(), documentDescriptor);

						IDocumentService documentService = officeApplication.getDocumentService();
						if (documentService == null)	System.out.println("LoadDocumentOperation: officeApplication.getDocumentService()==null");
						else							System.out.println("LoadDocumentOperation: officeApplication.getDocumentService()="
																			+ officeApplication.getDocumentService());
						document = documentService.loadDocument(frame, url.toString(), documentDescriptor);

						System.out.println("LoadDocumentOperation: POST");

						// Das folgende (showElement zumindest) wirkt von weiter
						// oben NICHT bis nach dem loadDocument() (NOCHMAL
						// getestet, stimmt!):
						layoutManager.hideAll(); // 201202230412js: This becomes
													// effective immediately.
													// Schützt gegen offenen
													// URL_TOOLBAR,
													// aber leider NICHT gegen
													// offenen F11
													// Formatvorlagen Dialog
						layoutManager.showElement(ILayoutManager.URL_MENUBAR); // ok
																				// und
																				// sinnvoll
						layoutManager.showElement(ILayoutManager.URL_STATUSBAR); // ok
																					// und
																					// sinnvoll
						// layoutManager.showElement(ILayoutManager.URL_TOOLBAR);
						// //führt zum Stehenbleiben
						layoutManager.showElement(ILayoutManager.URL_TOOLBAR_STANDARDBAR); // ok
						// layoutManager.showElement(ILayoutManager.URL_TOOLBAR_ALIGNMENTBAR);
						// //ok, aber nicht für text sondern für andere objekte
						// -> eher unnötig
						// layoutManager.showElement(ILayoutManager.URL_TOOLBAR_FONTWORKOBJECTBAR);
						// //führt zum Stehenbleiben
						// layoutManager.showElement(ILayoutManager.URL_TOOLBAR_FONTWORKSHAPETYPES);
						// //ok, aber recht leer?
						// layoutManager.showElement(ILayoutManager.URL_TOOLBAR_FORMATOBJECTBAR);
						// //ok, aber recht leer?

						// layoutManager.showElement(ILayoutManager.URL_TOOLBAR_PREVIEWBAR);
						// //ok, aber dieser und/oder viewerbar in standardbar
						// enthalten
						layoutManager.showElement(ILayoutManager.URL_TOOLBAR_TEXTOBJECTBAR); // ok
																								// und
																								// relativ
																								// sinnvoll
						// layoutManager.showElement(ILayoutManager.URL_TOOLBAR_VIEWERBAR);
						// //ok, aber dieser und/oder previewbar in standardbar
						// enthalten

						xLayoutManager.dockAllWindows(com.sun.star.ui.UIElementType.FLOATINGWINDOW);
						xLayoutManager.dockAllWindows(com.sun.star.ui.UIElementType.TOOLBAR);
						xLayoutManager.dockAllWindows(com.sun.star.ui.UIElementType.MENUBAR);
						xLayoutManager.dockAllWindows(com.sun.star.ui.UIElementType.POPUPMENU);
						xLayoutManager.dockAllWindows(com.sun.star.ui.UIElementType.STATUSBAR);
						xLayoutManager.dockAllWindows(com.sun.star.ui.UIElementType.PROGRESSBAR);
						xLayoutManager.dockAllWindows(com.sun.star.ui.UIElementType.UNKNOWN);

					} else
						document = officeApplication.getDocumentService().loadDocument(url.toString(), documentDescriptor);

					System.out.println("LoadDocumentOperation: InternalThread: run: Status after loadDocument()...");
					if (document == null)
						System.out.println("LoadDocumentOperation: InternalThread: run: document==null");
					else
						System.out.println("LoadDocumentOperation: InternalThread: run: document=" + document.toString());

				}
				System.out.println("LoadDocumentOperation: InternalThread: run: ...work has been done.");
				done = true;
			} catch (Exception exception) {
				System.out.println("LoadDocumentOperation: InternalThread: run: WARNING: Caught exception!");
				this.exception = exception;
			} catch (ThreadDeath threadDeath) {
				// do not consume
				System.out.println("LoadDocumentOperation: InternalThread: run: WARNING: Caught threadDeath!");
			}
			System.out.println("LoadDocumentOperation: InternalThread: run: ends");
		}

		// ----------------------------------------------------------------------------
		/**
		 * Returns exception. Returns null if no exception was thrown.
		 * 
		 * @return exception - returns null if no exception was thrown
		 * 
		 * @author Andreas Br�ker
		 */
		public Exception getException() {
			return exception;
		}

		// ----------------------------------------------------------------------------
		/**
		 * Returns loaded document.
		 * 
		 * @return loaded document
		 * 
		 * @author Andreas Br�ker
		 */
		public IDocument getDocument() {
			return document;
		}

		// ----------------------------------------------------------------------------
		/**
		 * Returns information whether the thread has finished his work.
		 * 
		 * @return information whether the thread has finished his work
		 * 
		 * @author Andreas Br�ker
		 */
		public boolean done() {
			if (exception != null)
				return true;
			return done;
		}
		// ----------------------------------------------------------------------------
	}

	// ----------------------------------------------------------------------------

	// ----------------------------------------------------------------------------
	/**
	 * Constructs new LoadDocumentOperation.
	 * 
	 * @param officeApplication
	 *            office application to be use
	 * @param frame
	 *            frame to be used
	 * @param url
	 *            URL of the document
	 * @param documentDescriptor
	 *            document descriptor to be used (can be null)
	 * 
	 * @throws IllegalArgumentException
	 *             if the submitted office application, frame or URL is not
	 *             valid
	 * 
	 * @author Andreas Br�ker
	 */
	public LoadDocumentOperation(IOfficeApplication officeApplication, IFrame frame, URL url, IDocumentDescriptor documentDescriptor)
			throws IllegalArgumentException {
		this(null, officeApplication, frame, url, documentDescriptor);
	}

	// ----------------------------------------------------------------------------
	/**
	 * Constructs new LoadDocumentOperation.
	 * 
	 * @param documentType
	 *            document type to be loaded (can be null)
	 * @param officeApplication
	 *            office application to be use
	 * @param frame
	 *            frame to be used
	 * @param url
	 *            URL of the document
	 * @param documentDescriptor
	 *            document descriptor to be used (can be null)
	 * 
	 * @author Andreas Br�ker
	 */
	public LoadDocumentOperation(String documentType, IOfficeApplication officeApplication, IFrame frame, URL url,
			IDocumentDescriptor documentDescriptor) {

		assert officeApplication != null;

		this.officeApplication = officeApplication;
		this.frame = frame;
		this.url = url;

		this.documentDescriptor = documentDescriptor;
		this.documentType = documentType;
	}

	// ----------------------------------------------------------------------------
	/**
	 * Constructs new LoadDocumentOperation.
	 * 
	 * @param documentType
	 *            document type to be loaded (can be null)
	 * @param officeApplication
	 *            office application to be use
	 * @param url
	 *            URL of the document
	 * @param documentDescriptor
	 *            document descriptor to be used (can be null)
	 * 
	 * @author Andreas Br�ker
	 */
	public LoadDocumentOperation(String documentType, IOfficeApplication officeApplication, URL url, IDocumentDescriptor documentDescriptor) {

		assert officeApplication != null;

		this.officeApplication = officeApplication;
		this.url = url;

		this.documentDescriptor = documentDescriptor;
		this.documentType = documentType;
	}

	// ----------------------------------------------------------------------------
	/**
	 * Constructs new LoadDocumentOperation.
	 * 
	 * @param documentType
	 *            document type to be loaded (can be null)
	 * @param officeApplication
	 *            office application to be use
	 * @param frame
	 *            frame to be used
	 * @param inputStream
	 *            input stream to be used
	 * @param documentDescriptor
	 *            document descriptor to be used (can be null)
	 * 
	 * @author Andreas Br�ker
	 * @date 06.07.2006
	 */
	public LoadDocumentOperation(String documentType, IOfficeApplication officeApplication, IFrame frame, InputStream inputStream,
			IDocumentDescriptor documentDescriptor) {
		assert officeApplication != null;
		assert inputStream != null;

		this.officeApplication = officeApplication;
		this.frame = frame;
		this.inputStream = inputStream;

		this.documentDescriptor = documentDescriptor;
		this.documentType = documentType;
	}

	// ----------------------------------------------------------------------------
	/**
	 * Constructs new LoadDocumentOperation.
	 * 
	 * @param officeApplication
	 *            office application to be use
	 * @param inputStream
	 *            input stream to be used
	 * @param documentDescriptor
	 *            document descriptor to be used (can be null)
	 * 
	 * @author Andreas Br�ker
	 * @date 10.07.2006
	 */
	public LoadDocumentOperation(IOfficeApplication officeApplication, InputStream inputStream, IDocumentDescriptor documentDescriptor) {
		this(null, officeApplication, null, inputStream, documentDescriptor);
	}

	// ----------------------------------------------------------------------------
	/**
	 * Sets information whether this operation is a sub task.
	 * 
	 * @param isSubTask
	 *            information whether this operation is a sub task
	 * 
	 * @author Andreas Br�ker
	 */
	public void setIsSubTask(boolean isSubTask) {
		this.isSubTask = isSubTask;
	}

	// ----------------------------------------------------------------------------
	/**
	 * Sets information whether the progress monitor should be updated.
	 * 
	 * @param updateProgressMonitor
	 *            information whether the progress monitor should be updated
	 * 
	 * @author Andreas Br�ker
	 * @date 10.07.2006
	 */
	public void setUpdateProgressMonitor(boolean updateProgressMonitor) {
		this.updateProgressMonitor = updateProgressMonitor;
	}

	// ----------------------------------------------------------------------------
	/**
	 * Sets information whether the document should be loaded by an input
	 * stream.
	 * 
	 * @param useStream
	 *            information whether the document should be loaded by an input
	 *            stream
	 * 
	 * @author Andreas Br�ker
	 * @date 13.06.2006
	 */
	public void setUseStream(boolean useStream) {
		this.useStream = useStream;
	}

	// ----------------------------------------------------------------------------
	/**
	 * Returns exception. Returns null if no exception was thrown.
	 * 
	 * @return exception - returns null if no exception was thrown
	 * 
	 * @author Andreas Br�ker
	 */
	public Exception getException() {
		return internalThread.getException();
	}

	// ----------------------------------------------------------------------------
	/**
	 * Returns loaded document. Returns null if no document is available. If a
	 * document type was submitted - only a document will be returned if the
	 * type of the loaded document matches the submitted type.
	 * 
	 * @return loaded document or null if no document is available
	 * 
	 * @author Andreas Br�ker
	 */
	public IDocument getDocument() {
		if (documentType == null)
			return internalThread.getDocument();
		else {
			IDocument document = internalThread.getDocument();
			if (document != null) {
				if (document.getDocumentType().equals(documentType))
					return document;
			}
			return null;
		}
	}

	// ----------------------------------------------------------------------------
	/**
	 * Runs this operation. Progress should be reported to the given progress
	 * monitor. This method is usually invoked by an
	 * <code>IRunnableContext</code>'s <code>run</code> method, which supplies
	 * the progress monitor. A request to cancel the operation should be honored
	 * and acknowledged by throwing <code>InterruptedException</code>.
	 * 
	 * @param progressMonitor
	 *            the progress monitor to use to display progress and receive
	 *            requests for cancelation
	 * 
	 * @exception InvocationTargetException
	 *                if the run method must propagate a checked exception, it
	 *                should wrap it inside an
	 *                <code>InvocationTargetException</code>; runtime exceptions
	 *                are automatically wrapped in an
	 *                <code>InvocationTargetException</code> by the calling
	 *                context
	 * @exception InterruptedException
	 *                if the operation detects a request to cancel, using
	 *                <code>IProgressMonitor.isCanceled()</code>, it should exit
	 *                by throwing <code>InterruptedException</code>
	 * 
     * Joerg Sigle added code for progress monitoring printlns, to examine the stability problems observed in Elexis,
     * and to implement a watchdog timer that will interrupt a hanging loop after a certain time.
     * Empirically, the call to internalThread.destroy(); apparently hits the sweet spot. 
     * It changes the user experience of a completely stalled Elexis/OpenOffice/LibreOffice program
     * into one that may merely be paused for several 10 seconds, and then continue - just as it would be expected.
     * Even when I opened a document with many interactive elements and lots of macro code which would
     * require more than one minute in LibreOffice (for whatever reason, same time if opened outside Elexis),
     * the watchdog would kick in, but the document would still be loaded correctly and work correctly thereafter.
     * So it seems that the watchdog improves the situation and, give the duration I've told it to wait,
     * does not have any disadvantageous effects.
     * In the future, however, I would suggest to review what is actually happening in the interface
     * to greater depth, i.e. down into the OpenOffice/LibreOffice libraries. I haven't done this so far,
     * as I did all the work without funding, with other projects pressing, so I stopped - for now -
     * as son as I saw I had to add the source code of the Office packages into my development system,
     * and my solution found until now would solve the problem in praxi (I have given it several tests).
     * Further details in the comprehensive comments I added below and elsewhere in the library.
     * 
     * @author Joerg Sigle
     * @date 25.02.2012
     * 
	 * @author Andreas Br�ker
	 */
	public void run(IProgressMonitor progressMonitor) throws InvocationTargetException, InterruptedException {
		System.out.println("LoadDocumentOperation: run begins");

		System.out.println("LoadDocumentOperation: run 1");
		internalThread = new InternalThread();
		if (isSubTask) {
			if (updateProgressMonitor)
				progressMonitor.subTask(Messages.LoadDocumentOperation_monitor_loading_document);
		} else {
			if (updateProgressMonitor)
				progressMonitor.beginTask(Messages.LoadDocumentOperation_monitor_loading_document, 50);
		}

		System.out.println("LoadDocumentOperation: run 2");
		internalThread.start();
		System.out.println("LoadDocumentOperation: run 2.1");
		int cyclesWaitedForInternalThreadToComplete=0;
		while (!internalThread.done()) {
			System.out.println("LoadDocumentOperation: run 2.2 sleep(500) begins");
			Thread.sleep(500);
			System.out.println("LoadDocumentOperation: run 2.3 sleep(500) ends");
			if (!isSubTask)
				progressMonitor.worked(1);
			System.out.println("LoadDocumentOperation: run 2.4");
			if (progressMonitor.isCanceled()) {
				/**
				 * This method is deprecated, but ...
				 */
				try {
					System.out.println("LoadDocumentOperation: run 2.5");
					internalThread.stop();
				} catch (Throwable throwable) {
					// do not consume - ThreadDeath
					System.out.println("LoadDocumentOperation: run 2.6");
				}
				System.out.println("LoadDocumentOperation: run 2.7");
				progressMonitor.done();
				System.out.println("LoadDocumentOperation: run 2.8 about to throw InterruptException()...");
				throw new InterruptedException(Messages.LoadDocumentOperation_exception_message_operation_interrupted);
			}
			//201202252105js:
			//Now limit the time we wait for the office application to start.
			//As it may hang (when a floating dialog window like the F11 Formatvorlagen dialog, or table frame properties)
			//is open in OpenOffice/LibreOffice, we will now set a limit for the time we want to wait,
			//and kill the internalThread if that is exceeded.
			//NOTE: This hanging has absolutely NOTHING to do with the removed sleep(3000) below.
			//It has occured forever since oo 2.0.3 in all variants of the noatext interface I now,
			//even after my update to noa 2.2.3 and noa-libre. See my external logs and notes in the code
			//for the results of my debugging.
			//NOTE: On an Intel Core i7 vpro notebook, loading a document takes only 1..2 seconds when it works.
			//To accomodate for significantly slower infrastructure, I'll wait for 20 seconds now before killing the internalThread.
			//NOTE: internalThread is the one that should take care of the loading, which it normally does,
			//after having jumped through zillions of references, methods, procedures etc. within the Java and StarOffice library jungle.
			//NOTE: The NOAText_js and NOAText_jsl implementations of the NOAText plugin, take numerous preconditions
			//to completely avoid user generated situations where internalThread would be likely to never come to an end.
			//Especially, they disable a number of menu commands that could open probably persistent dialogs or toolbars,
			//which are not probably needed by users within Elexis. Moreover, they try to hide all toolbars and only
			//display the probably needed ones directly after loading the document (which is the earliest point in time
			//when it is beneficial, according to practical tests). But still, some dialogs can probably not be guaranteed
			//to never be required, and never persist, so its these remaining few against which the following timeout handler
			//shall finally help.
			//NOTE, IMPORTANT: It would also be a good precondition to "fixate" the toolbars within OpenOffice.
			//I just now remember this was a measure that I had originally taken on all our workstations, and it also appeared to help.
			//NOTE: Apparently - as demonstrated in the Mustermann letter with the table format dialog, the following beneficial thing happens:
			//If (!) the user detaches the table format dialog from the toolbars, that dialog becomes floating.
			//If he then closes the Briefe view, the floating dialog state is saved.
			//If he then re-loads the same Letter, the open floating dialog would stall the loader thread (for any reason whatsoever)
			//After some time, the limit below kicks in and the internalThread is destroyed.
			//The previously stalled OpenOffice, however, appears in the frame, and most probably...
			//my precautions shown somewhere else do immediately reconfigure the toolbars and cause the floating dialog to become docked!
			//When the user closes the window again, the new state is saved; and in future sessions, the dialog will not be floating any more
			//even when the same document is re-loaded. :-)
			//(Until the user decides to actively pull it out of the dock again, which should hardly ever happen.)
			//So only in very rare occasions, this very rude timout driven stall recovery needs to kick in,
			//and its benefits extend beyond resolving the acute situation. For a workaround, I think that's rather nice.
			//201202252154js
			cyclesWaitedForInternalThreadToComplete=cyclesWaitedForInternalThreadToComplete+1;	//Each cycle currently sleeps for 500ms
			if (cyclesWaitedForInternalThreadToComplete>40) {
				cyclesWaitedForInternalThreadToComplete=0; //Paranoia: Damit der nächste Thread.destroy() ggf. nicht schon in 0.5 Sek zuschlagen kann...
				System.out.println("LoadDocumentOperation: ERROR: internalThread timed out; will now be destroyed.");
				System.out.println("WARNING: This functionality should NOT be used except for debugging - see intenalThread.destroy() help .");
				System.out.println("WARNING: Destroying the internalThread is not guaranteed to return Elexis into a responsive state.");
				System.out.println("WARNING: Neither can we guarantee that the document that was about to be opened will appear correctly.");
				System.out.println("WARNING: That said, however, it DID succeed in some tests (4/5), and made the desired document");
				System.out.println("WARNING: appear with LibreOffice in its desired frame, and even the problematic floating dialog windows");
				System.out.println("WARNING: in addition to that - for either the table properties, or the externally opened F11 Formatvorlagen.");
				System.out.println("WARNING: 201202252133js");
				internalThread.destroy();	//this makes progressMonitor detect the end of the thread at least,
											//but it is deprecated and very error prone. So this should NOT be used except for debugging.
				//internalThread.stop();	//this does not stop progressMonitor from cycling on and on.
				Thread.sleep(500);			//Nach dem destroy des anderen Threads lasse ich hier noch eine halbe Sekunde Zeit.
			}
			
		}

		System.out.println("LoadDocumentOperation: run 3 TO DO: Thread.sleep(3000) [WARNING: REMOVED SLEEP()], otherwise OO 3.x might crash (says comment in code)");
		// sleep here for a while otherwise OOo 3.x might crash
		//Thread.sleep(3000);
		System.out.println("LoadDocumentOperation: run 4 sleep ends");

		if (!isSubTask)
			progressMonitor.done();
		System.out.println("LoadDocumentOperation: run 5 done");
	}
	// ----------------------------------------------------------------------------

}