/****************************************************************************
 * ubion.ORS - The Open Report Suite                                        *
 *                                                                          *
 * ------------------------------------------------------------------------ *
 *                                                                          *
 * Subproject: Office Editor Core                                           *
 *                                                                          *
 *                                                                          *
 * The Contents of this file are made available subject to                  *
 * the terms of GNU Lesser General Public License Version 2.1.              *
 *                                                                          * 
 * GNU Lesser General Public License Version 2.1                            *
 * ======================================================================== *
 * Copyright 2003-2005 by IOn AG                                            *
 *                                                                          *
 * This library is free software; you can redistribute it and/or            *
 * modify it under the terms of the GNU Lesser General Public               *
 * License version 2.1, as published by the Free Software Foundation.       *
 *                                                                          *
 * This library is distributed in the hope that it will be useful,          *
 * but WITHOUT ANY WARRANTY; without even the implied warranty of           *
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU        *
 * Lesser General Public License for more details.                          *
 *                                                                          *
 * You should have received a copy of the GNU Lesser General Public         *
 * License along with this library; if not, write to the Free Software      *
 * Foundation, Inc., 59 Temple Place, Suite 330, Boston,                    *
 * MA  02111-1307  USA                                                      *
 *                                                                          *
 * Contact us:                                                              *
 *  http://www.ion.ag                                                       *
 *  info@ion.ag                                                             *
 *                                                                          *
 ****************************************************************************/
 
/*
 * Last changes made by $Author: markus $, $Date: 2008-09-30 19:59:51 +0200 (Di, 30 Sep 2008) $
 */
package ag.ion.bion.workbench.office.editor.core;

import ag.ion.bion.officelayer.application.IOfficeApplication;
import ag.ion.bion.officelayer.application.OfficeApplicationRuntime;
import ag.ion.bion.officelayer.runtime.IOfficeProgressMonitor;

import org.eclipse.core.runtime.FileLocator;
import org.eclipse.core.runtime.IStatus;
import org.eclipse.core.runtime.Platform;
import org.eclipse.core.runtime.Plugin;
import org.eclipse.core.runtime.Status;

import org.osgi.framework.BundleContext;

import java.io.File;

import java.net.URL;

import java.util.HashMap;
import java.util.ResourceBundle;
import java.util.MissingResourceException;

import java.awt.Frame;

/**
 * The main plugin class to be used in the desktop.
 * 
 * @author Andreas Bröker
 * @version $Revision: 11647 $
 */
public class EditorCorePlugin extends Plugin {
  
  /** ID of the plugin. */
  public static final String PLUGIN_ID = "ag.ion.bion.workbench.office.editor.core";
  
  //The shared instance.
  private static EditorCorePlugin plugin;
  //Resource bundle.
  private ResourceBundle resourceBundle;
  
  private IOfficeApplication localOfficeApplication = null;
  
  private String librariesLocation = null;
	
  //----------------------------------------------------------------------------
  /**
   * The constructor.
   * 
   * @author Andreas Bröker
   */
  public EditorCorePlugin() {
    super();
		plugin = this;
		try {
		  resourceBundle = ResourceBundle.getBundle("ag.ion.bion.workbench.office.editor.core.CorePluginResources");
		} 
    catch (MissingResourceException missingResourceException) {
      resourceBundle = null;
    }
	}
  //----------------------------------------------------------------------------
  /**
   * This method is called upon plug-in activation.
   * 
   * @param context context to be used
   * 
   * @throws Exception if the bundle can not be started
   * 
   * @author Andreas Bröker
   */
  public void start(BundleContext context) throws Exception {
    super.start(context);
    System.setProperty(IOfficeApplication.NOA_NATIVE_LIB_PATH,getLibrariesLocation());
    /**
     * Workaround in order to integrate the OpenOffice.org window into a AWT frame
     * on Linux based systems. 
     */
    try {
      new Frame();
    }
    catch(Throwable throwable) {
      //only occurs in headless mode, where it doesn't matter
    }
  }
  //----------------------------------------------------------------------------
  /**
   * This method is called when the plug-in is stopped.
   * 
   * @param context context to be used
   * 
   * @throws Exception if the bundle can not be stopped
   * 
   * @author Andreas Bröker
   */
  public void stop(BundleContext context) throws Exception {
    super.stop(context);
  }
  //----------------------------------------------------------------------------
  /**
   * Returns the shared instance.
   * 
   * @return shared instance
   * 
   * @author Andreas Bröker
   */
  public static EditorCorePlugin getDefault() {
    return plugin;
  }
  //----------------------------------------------------------------------------
  /**
   * Returns the string from the plugin's resource bundle,
   * or 'key' if not found.
   * 
   * @param key key to be used
   * 
   * @return string from the plugin's resource bundle,
   * or 'key' if not found
   * 
   * @author Andreas Bröker
   */
  public static String getResourceString(String key) {
    ResourceBundle bundle = EditorCorePlugin.getDefault().getResourceBundle();
		try {
		  return (bundle != null) ? bundle.getString(key) : key;
		} 
    catch (MissingResourceException missingResourceException) {
      return key;
    }
  }
  //----------------------------------------------------------------------------
  /**
   *  Returns the plugin's resource bundle.
   * 
   * @return plugin's resource bundle
   * 
   * @author Andreas Bröker
   */
  public ResourceBundle getResourceBundle() {
    return resourceBundle;
  }
  //----------------------------------------------------------------------------
  /**
   * Returns local office application. The instance of the application
   * will be managed by this plugin.
   * 
   * @return local office application
   * 
   * @author Andreas Bröker
   */
  public synchronized IOfficeApplication getManagedLocalOfficeApplication() {
    if(localOfficeApplication == null) {
      HashMap configuration = new HashMap(1);
      configuration.put(IOfficeApplication.APPLICATION_TYPE_KEY, IOfficeApplication.LOCAL_APPLICATION);
      try {
        localOfficeApplication = OfficeApplicationRuntime.getApplication(configuration);
      }
      catch(Throwable throwable) {
        //can not be - this code must work
        Platform.getLog(getBundle()).log(new Status(IStatus.ERROR, EditorCorePlugin.PLUGIN_ID,
            IStatus.ERROR, throwable.getMessage(), throwable));
      }
    }
    return localOfficeApplication;
  }
  //----------------------------------------------------------------------------
  /**
   * Returns location of the libraries of the plugin. Returns null if the location
   * can not be provided.
   * 
   * @return location of the libraries of the plugin or null if the location
   * can not be provided
   * 
   * @author Andreas Bröker
   */
  public String getLibrariesLocation() {
    if(librariesLocation == null) {
      try {
        URL url = Platform.getBundle("ag.ion.noa").getEntry("/");
        url  = FileLocator.toFileURL(url);
        String bundleLocation = url.getPath();
        File file = new File(bundleLocation);
        bundleLocation = file.getAbsolutePath();
        bundleLocation = bundleLocation.replace('/', File.separatorChar) + File.separator + "lib";
        librariesLocation = bundleLocation;
      }
      catch(Throwable throwable) {
        return null;
      }
    }
    return librariesLocation;
  }  
  //----------------------------------------------------------------------------
  
}