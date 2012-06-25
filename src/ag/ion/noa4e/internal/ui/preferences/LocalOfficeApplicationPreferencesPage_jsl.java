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
 * Portions Copyright 2007 by Gerry Weirich (Only for a different branch    *
 *  producing his NOAText 1.4.1, not directly used to generate this file.)  *
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
 *  http://www.elexis.ch                                                    *
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
 * Possibly, this preference page should get a new and unique PAGE_ID string.
 * Possibly, this version of the library should get a new revision number, currently used is 11685.
 * Clarificytion of LGPL vs. GPL vs. Eclipse licensing issues.
 ****************************************************************************/

/*
 * Last changes made by $Author: jsigle $, $Date: 2012-06-23 14:38:00 +0100 (Su, 23 Jun 2012) $
 */
package ag.ion.noa4e.internal.ui.preferences;

import java.util.Arrays;
import java.util.TreeSet;

import org.eclipse.jface.dialogs.MessageDialog;
import org.eclipse.jface.preference.IPreferenceStore;
import org.eclipse.jface.preference.PreferencePage;
import org.eclipse.jface.resource.JFaceResources;
import org.eclipse.jface.viewers.TableLayout;
import org.eclipse.jface.window.Window;
import org.eclipse.jface.wizard.WizardDialog;
import org.eclipse.swt.SWT;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.layout.GridData;
import org.eclipse.swt.layout.GridLayout;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.widgets.Composite;
import org.eclipse.swt.widgets.Control;
import org.eclipse.swt.widgets.Label;
import org.eclipse.swt.widgets.Link;
import org.eclipse.swt.widgets.Table;
import org.eclipse.swt.widgets.TableColumn;
import org.eclipse.swt.widgets.TableItem;
import org.eclipse.swt.widgets.Text;
import org.eclipse.ui.IWorkbench;
import org.eclipse.ui.IWorkbenchPreferencePage;
import org.eclipse.ui.forms.widgets.FormToolkit;

import ch.elexis.Hub;
import ch.elexis.preferences.PreferenceConstants;
import ch.elexis.preferences.SettingsPreferenceStore;

import ag.ion.bion.officelayer.application.IApplicationAssistant;
import ag.ion.bion.officelayer.application.IApplicationProperties;
import ag.ion.bion.officelayer.application.ILazyApplicationInfo;
import ag.ion.bion.officelayer.application.OfficeApplicationRuntime;
import ag.ion.bion.workbench.office.editor.core.EditorCorePlugin;
import ag.ion.noa4e.ui.NOAUIPlugin;
import ag.ion.noa4e.ui.wizards.application.LocalApplicationWizard;

/**
 * Preferences page for local OpenOffice.org application - adopted for Elexis and NOAText_jsl.
 * 
 * @author Joerg Sigle
 * @author Andreas Br�ker
 * @author Markus Kr�ger
 * @version $Revision: 11685 $
 */
public class LocalOfficeApplicationPreferencesPage_jsl extends PreferencePage implements
    IWorkbenchPreferencePage {

  /** ID of the page. */
  public static final String PAGE_ID = "ag.ion.noa4e.ui.preferences.LocalOfficeApplicationPreferencePage_jsl"; //$NON-NLS-1$
  
  /**
   * @author Joerg Sigle
   * 
   * Adopted for Elexis by Joerg Sigle 02/2012, adding the following line.
   * Changes required because of different preference store layout in Elexis.
   * There are corresponding changes in:
   * LocalOfficeApplicationsPreferencesPage.java
   *   PREFS_PREVENT_TERMINATION
   *   initPreferenceValues()
   *   performOk()
   * NOAUIPlugin.java							
   *   PREFERENCE_OFFICE_HOME
   *   PREFERENCE_PREVENT_TERMINATION				
   *   internalStartApplication().
   */
  public static final String PREFS_PREVENT_TERMINATION	= "openoffice/preventTermination";

  private Text               textHome                   = null;
  private Button             buttonPreventTermination   = null;

  private Table              tableApplicationProperties = null;

  //----------------------------------------------------------------------------
  /**
   * Initializes this preference page for the given workbench.
   * 
   * @param workbench workbnech to be used
   * 
   * @author Andreas Br�ker
   */
  public void init(IWorkbench workbench) {
	  System.out.println("LOAPP: init");
		  setDescription(Messages.LocalOfficeApplicationPreferencesPage_description_configure_application);
  }

  //----------------------------------------------------------------------------
  /**
   * Creates and returns the SWT control for the customized body of this preference 
   * page under the given parent composite. 
   * 
   * @param parent the parent composite
   * 
   * @return constructed control
   * 
   * @author Andreas Br�ker
   * @author Markus Kr�ger
   */
  protected Control createContents(Composite parent) {
	System.out.println("LOAPP: createContents");
	FormToolkit formToolkit = NOAUIPlugin.getFormToolkit();
    Composite composite = new Composite(parent, SWT.NULL);

    GridLayout gridLayout = new GridLayout();
    gridLayout.numColumns = 3;
    composite.setLayout(gridLayout);

    Label labelHome = formToolkit.createLabel(composite,
        Messages.LocalOfficeApplicationPreferencesPage_label_application_home);
    labelHome.setBackground(composite.getBackground());

    textHome = formToolkit.createText(composite, ""); //$NON-NLS-1$
    textHome.setEditable(false);
    textHome.setFont(composite.getFont());
    GridData gridData = new GridData(SWT.FILL, SWT.NONE, true, false);
    textHome.setLayoutData(gridData);

    final Link linkDefine = new Link(composite, SWT.NONE);
    linkDefine.setText("<a>" + Messages.LocalOfficeApplicationPreferencesPage_link_define_text + "</a>"); //$NON-NLS-1$ //$NON-NLS-2$ //$NON-NLS-3$
    linkDefine.addSelectionListener(new SelectionAdapter() {
      public void widgetSelected(SelectionEvent selectionEvent) {
    	System.out.println("LOAPP: createContents: widgetSelected: 1 start");
    	LocalApplicationWizard localApplicationWizard = new LocalApplicationWizard();
    	System.out.println("LOAPP: createContents: widgetSelected: 2");
    	String oldHome = textHome.getText();
    	System.out.println("LOAPP: createContents: widgetSelected: 3");
    	if (oldHome.length() != 0)
          localApplicationWizard.setHomePath(oldHome);
    	System.out.println("LOAPP: createContents: widgetSelected: 4");
    	WizardDialog wizardDialog = new WizardDialog(linkDefine.getShell(), localApplicationWizard);
    	System.out.println("LOAPP: createContents: widgetSelected: 5");
    	if (wizardDialog.open() == Window.OK) {
    		System.out.println("LOAPP: createContents: widgetSelected: 6");
        	String home = localApplicationWizard.getSelectedHomePath();
        	System.out.println("LOAPP: createContents: widgetSelected: 7");
        	if (home != null)
        		textHome.setText(home);
          initApplicationProperties(tableApplicationProperties);
          System.out.println("LOAPP: createContents: widgetSelected: 8");
      	}
    	System.out.println("LOAPP: createContents: widgetSelected: 9 end");
      }
    });

    Label labelNull = formToolkit.createLabel(composite, ""); //$NON-NLS-1$
    gridData = new GridData();
    gridData.horizontalSpan = 3;
    labelNull.setLayoutData(gridData);

    Label labelProperties = formToolkit.createLabel(composite,
        Messages.LocalOfficeApplicationPreferencesPage_label_application_properties_text);
    labelProperties.setBackground(composite.getBackground());
    labelProperties.setFont(JFaceResources.getFontRegistry().getBold(labelProperties.getFont().toString()));
    gridData = new GridData(SWT.FILL, SWT.NONE, true, false);
    gridData.horizontalSpan = 3;
    labelProperties.setLayoutData(gridData);

    tableApplicationProperties = formToolkit.createTable(composite, SWT.READ_ONLY);
    gridData = new GridData(SWT.FILL, SWT.FILL, true, true);
    gridData.horizontalSpan = 3;
    int tableWidth = (int) (tableApplicationProperties.getDisplay().getClientArea().width * 0.3);
    gridData.widthHint = tableWidth;
    tableApplicationProperties.setLayoutData(gridData);

    TableLayout tableLayout = new TableLayout();
    tableApplicationProperties.setLayout(tableLayout);

    TableColumn columnProduct = new TableColumn(tableApplicationProperties, SWT.NONE);
    columnProduct.setText(Messages.LocalOfficeApplicationPreferencesPage_column_name_text);
    int columnProductWidth = (int) (tableWidth * 0.4);
    columnProduct.setWidth(columnProductWidth);

    TableColumn columnHome = new TableColumn(tableApplicationProperties, SWT.NONE);
    columnHome.setText(Messages.LocalOfficeApplicationPreferencesPage_column_value_text);
    columnHome.setWidth(tableWidth - columnProductWidth);

    tableApplicationProperties.setLinesVisible(true);
    tableApplicationProperties.setHeaderVisible(true);

    buttonPreventTermination = formToolkit.createButton(composite,
        Messages.LocalOfficeApplicationPreferencesPage_prevent_termination_lable,
        SWT.CHECK);
    buttonPreventTermination.setBackground(composite.getBackground());
    gridData = new GridData();
    gridData.horizontalSpan = 3;
    buttonPreventTermination.setLayoutData(gridData);

    formToolkit.paintBordersFor(composite);
    initPreferenceValues();
    initApplicationProperties(tableApplicationProperties);
    return composite;
  }

  //----------------------------------------------------------------------------
  /**
   * Notifies that the OK button of this page's container has been pressed. 
   * 
   * @return false to abort the container's OK processing and true to allow 
   * the OK to happen
   * 
   * @author Joerg Sigle
   * @author Gerry Weirich
   * @author Andreas Br�ker
   * @author Markus Kr�ger
   *
   * Adopted for Elexis by Joerg Sigle 02/2012, adding comments and monitoring output,
   * and reproducing the functionality of changes made by Gerry Weirich in 06/2007
   * for his NOAText plugin 1.4.1 to a file obtained from an older version of the ag.ion noa library.
   * 
   * Changes required because of different preference store layout in Elexis:
   * There are corresponding changes in:
   * LocalOfficeApplicationsPreferencesPage.java
   *   PREFS_PREVENT_TERMINATION
   *   initPreferenceValues()
   *   performOk()
   * NOAUIPlugin.java							
   *   PREFERENCE_OFFICE_HOME
   *   PREFERENCE_PREVENT_TERMINATION				
   *   internalStartApplication().
   */
  public boolean performOk() {
	System.out.println("LOAPP: performOK - Adopted to Elexis by GW/JS");
	System.out.println("LOAPP: allocating preferenceStore = new SettingsPreferenceStore(Hub.localCfg)");
	System.out.println("LOAPP: instead of using = NOAUIPlugin.getDefault().getPreferenceStore()");
	
	IPreferenceStore preferenceStore = new SettingsPreferenceStore(Hub.localCfg);
    preferenceStore.setValue(PREFS_PREVENT_TERMINATION, buttonPreventTermination.getSelection());
   
	//IPreferenceStore preferenceStore = NOAUIPlugin.getDefault().getPreferenceStore();
    //preferenceStore.setValue(NOAUIPlugin.PREFERENCE_PREVENT_TERMINATION,
    //    buttonPreventTermination.getSelection());

    String oldPath = preferenceStore.getString(PreferenceConstants.P_OOBASEDIR);
    preferenceStore.setValue(PreferenceConstants.P_OOBASEDIR, textHome.getText());

    //String oldPath = preferenceStore.getString(NOAUIPlugin.PREFERENCE_OFFICE_HOME);
    //preferenceStore.setValue(NOAUIPlugin.PREFERENCE_OFFICE_HOME, textHome.getText());

    System.out.println("LOAPP: Please note: There is a reference to NOAUIPlugin.getDefault()...");
    System.out.println("LOAPP: still left in this code; I (js) don't know whether this might be null and hence not work.");
        
    super.performOk();
    if (oldPath.length() != 0 || !oldPath.equals(textHome.getText())) {
      if (EditorCorePlugin.getDefault().getManagedLocalOfficeApplication().isActive()) {
        if (MessageDialog.openQuestion(getShell(),
            Messages.LocalOfficeApplicationPreferencesPage_dialog_restart_workbench_title,
            Messages.LocalOfficeApplicationPreferencesPage_dialog_restart_workbench_message))
          NOAUIPlugin.getDefault().getWorkbench().restart();
      }
    }
    return true;
  }

  //----------------------------------------------------------------------------
  /**
   * Inits application properties.
   * 
   * @param table table to be used
   * 
   * @author Andreas Br�ker
   */
  private void initApplicationProperties(Table table) {
	  System.out.println("LOAPP: initApplicationProperties");
	  try {
      TableItem[] tableItems = table.getItems();
      for (int i = 0, n = tableItems.length; i < n; i++) {
        tableItems[i].dispose();
      }

      IApplicationAssistant applicationAssistant = OfficeApplicationRuntime.getApplicationAssistant(EditorCorePlugin.getDefault().getLibrariesLocation());
      ILazyApplicationInfo applicationInfo = applicationAssistant.findLocalApplicationInfo(textHome.getText());
      if (applicationInfo != null) {
        IApplicationProperties applicationProperties = applicationInfo.getProperties();
        if (applicationProperties != null) {
          String[] names = applicationProperties.getPropertyNames();
          TreeSet treeSet = new TreeSet(Arrays.asList(names));
          names = (String[]) treeSet.toArray(new String[treeSet.size()]);
          for (int i = 0, n = names.length; i < n; i++) {
            String name = names[i];
            String value = applicationProperties.getPropertyValue(name);
            if (value != null && value.length() != 0) {
              TableItem tableItem = new TableItem(table, SWT.NONE);
              tableItem.setText(0, name);
              tableItem.setText(1, value);
            }
          }
        }
      }
    }
    catch (Throwable throwable) {
      //do not consume
    }
  }

  //----------------------------------------------------------------------------
  /**
   * Inits all preference values.
   * 
   * @author Joerg Sigle
   * @author Gerry Weirich
   * @author Andreas Br�ker
   * @author Markus Kr�ger
   *
   * Adopted for Elexis by Joerg Sigle 02/2012, adding comments and monitoring output,
   * and reproducing the functionality of changes made by Gerry Weirich in 06/2007
   * for his NOAText plugin 1.4.1 to a file obtained from an older version of the ag.ion noa library.
   * 
   * Changes required because of different preference store layout in Elexis.
   * There are corresponding changes in:
   * LocalOfficeApplicationsPreferencesPage.java
   *   PREFS_PREVENT_TERMINATION
   *   initPreferenceValues()
   *   performOk()
   * NOAUIPlugin.java			
   *   PREFERENCE_OFFICE_HOME
   *   PREFERENCE_PREVENT_TERMINATION				
   *   internalStartApplication().
   */
  private void initPreferenceValues() {
	System.out.println("LOAPP: initPreferenceValues - adopted for Elexis and NOAText_jsl by GW/JS");
	System.out.println("LOAPP: allocating preferenceStore = new SettingsPreferenceStore(Hub.localCfg)");
	System.out.println("LOAPP: instead of using = NOAUIPlugin.getDefault().getPreferenceStore()");
	
	IPreferenceStore preferenceStore=new SettingsPreferenceStore(Hub.localCfg);
	String officeHomePath=preferenceStore.getString(PreferenceConstants.P_OOBASEDIR);
	boolean preventTermination=preferenceStore.getBoolean(PREFS_PREVENT_TERMINATION);
	  
	//IPreferenceStore preferenceStore = NOAUIPlugin.getDefault().getPreferenceStore();
    //String officeHomePath = preferenceStore.getString(NOAUIPlugin.PREFERENCE_OFFICE_HOME);
    //boolean preventTermination = preferenceStore.getBoolean(NOAUIPlugin.PREFERENCE_PREVENT_TERMINATION);

    textHome.setText(officeHomePath);
    buttonPreventTermination.setSelection(preventTermination);
  }

  //----------------------------------------------------------------------------
  /**
   * Returns information whether this preferences page is valid.
   * 
   * @return information whether this preferences page is valid
   * 
   * @author Andreas Br�ker
   */
  public boolean isValid() {
	  System.out.println("LOAPP: isValid - always just returns true");
	  return true;
  }
  //----------------------------------------------------------------------------

}