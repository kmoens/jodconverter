//
// JODConverter - Java OpenDocument Converter
// Copyright 2004-2012 Mirko Nasato and contributors
//
// JODConverter is Open Source software, you can redistribute it and/or
// modify it under either (at your option) of the following licenses
//
// 1. The GNU Lesser General Public License v3 (or later)
// -> http://www.gnu.org/licenses/lgpl-3.0.txt
// 2. The Apache License, Version 2.0
// -> http://www.apache.org/licenses/LICENSE-2.0.txt
//
package org.artofsolving.jodconverter.util;

import java.io.File;
import java.util.Arrays;
import java.util.List;

public class PlatformUtils {

    private static String OS_NAME = System.getProperty("os.name").toLowerCase();

    private static final String WINDOWS = "windows";

    private static final String MAC = "mac";

    private static final String LINUX = "linux";

    private static final String[] LINUX_OO_HOME_PATHS = {
            "/usr/lib/libreoffice",
            "/opt/libreoffice4.0",
            "/usr/lib/openoffice",
            "/usr/lib/openoffice.org",
            "/usr/lib/openoffice.org3",
            "/opt/openoffice.org3",
            "/opt/libreoffice",
            "/usr/lib/ooo"
    };

    private static final String[] MAC_OO_HOME_PATHS = {
            "/Applications/LibreOffice.app/Contents",
            "/Applications/OpenOffice.org.app/Contents"
    };

    private static final String[] WINDOWS_OO_HOME_PATHS = {
            System.getenv("ProgramFiles") + File.separator + "LibreOffice 5.5",
            System.getenv("ProgramFiles") + File.separator + "LibreOffice 5.4",
            System.getenv("ProgramFiles") + File.separator + "LibreOffice 5.3",
            System.getenv("ProgramFiles") + File.separator + "LibreOffice 5.2",
            System.getenv("ProgramFiles") + File.separator + "LibreOffice 5.1",
            System.getenv("ProgramFiles") + File.separator + "LibreOffice 5.0",
            System.getenv("ProgramFiles") + File.separator + "LibreOffice 5",
            System.getenv("ProgramFiles") + File.separator + "LibreOffice 4.5",
            System.getenv("ProgramFiles") + File.separator + "LibreOffice 4.4",
            System.getenv("ProgramFiles") + File.separator + "LibreOffice 4.3",
            System.getenv("ProgramFiles") + File.separator + "LibreOffice 4.2",
            System.getenv("ProgramFiles") + File.separator + "LibreOffice 4.1",
            System.getenv("ProgramFiles") + File.separator + "LibreOffice 4.0",
            System.getenv("ProgramFiles") + File.separator + "LibreOffice 4",
            System.getenv("ProgramFiles") + File.separator + "LibreOffice 3.9",
            System.getenv("ProgramFiles") + File.separator + "LibreOffice 3.8",
            System.getenv("ProgramFiles") + File.separator + "LibreOffice 3.7",
            System.getenv("ProgramFiles") + File.separator + "LibreOffice 3.6",
            System.getenv("ProgramFiles") + File.separator + "LibreOffice 3.5",
            System.getenv("ProgramFiles") + File.separator + "LibreOffice 3.4",
            System.getenv("ProgramFiles") + File.separator + "LibreOffice 3",
            System.getenv("ProgramFiles(x86)") + File.separator + "LibreOffice 5.5",
            System.getenv("ProgramFiles(x86)") + File.separator + "LibreOffice 5.4",
            System.getenv("ProgramFiles(x86)") + File.separator + "LibreOffice 5.3",
            System.getenv("ProgramFiles(x86)") + File.separator + "LibreOffice 5.2",
            System.getenv("ProgramFiles(x86)") + File.separator + "LibreOffice 5.1",
            System.getenv("ProgramFiles(x86)") + File.separator + "LibreOffice 5.0",
            System.getenv("ProgramFiles(x86)") + File.separator + "LibreOffice 5",
            System.getenv("ProgramFiles(x86)") + File.separator + "LibreOffice 4.5",
            System.getenv("ProgramFiles(x86)") + File.separator + "LibreOffice 4.4",
            System.getenv("ProgramFiles(x86)") + File.separator + "LibreOffice 4.3",
            System.getenv("ProgramFiles(x86)") + File.separator + "LibreOffice 4.2",
            System.getenv("ProgramFiles(x86)") + File.separator + "LibreOffice 4.1",
            System.getenv("ProgramFiles(x86)") + File.separator + "LibreOffice 4.0",
            System.getenv("ProgramFiles(x86)") + File.separator + "LibreOffice 4",
            System.getenv("ProgramFiles(x86)") + File.separator + "LibreOffice 3.9",
            System.getenv("ProgramFiles(x86)") + File.separator + "LibreOffice 3.8",
            System.getenv("ProgramFiles(x86)") + File.separator + "LibreOffice 3.7",
            System.getenv("ProgramFiles(x86)") + File.separator + "LibreOffice 3.6",
            System.getenv("ProgramFiles(x86)") + File.separator + "LibreOffice 3.5",
            System.getenv("ProgramFiles(x86)") + File.separator + "LibreOffice 3.4",
            System.getenv("ProgramFiles(x86)") + File.separator + "LibreOffice 3",
            System.getenv("ProgramFiles") + File.separator + "OpenOffice.org 3",
            System.getenv("ProgramFiles(x86)") + File.separator + "OpenOffice.org 3"
    };

    private static String officeHome = null;

    private PlatformUtils() {
        throw new AssertionError("utility class must not be instantiated");
    }

    public static boolean isLinux() {
        return OS_NAME.startsWith(LINUX);
    }

    public static boolean isMac() {
        return OS_NAME.startsWith(MAC);
    }

    public static boolean isWindows() {
        return OS_NAME.startsWith(WINDOWS);
    }

    /**
     * Search for OpenOffice or LibreOffice on default paths.
     *
     * @return path to Office home or an empty String if not found.
     */
    public static String findOfficeHome() {
        if (officeHome == null) {
            String[] homeList = new String[0];
            if (isLinux()) {
                homeList = LINUX_OO_HOME_PATHS;
            } else if (isMac()) {
                homeList = MAC_OO_HOME_PATHS;
            } else if (isWindows()) {
                homeList = WINDOWS_OO_HOME_PATHS;
            }
            officeHome = searchExistingfile(Arrays.asList(homeList));
        }
        return officeHome;
    }

    public static void set(final String osName) {
        OS_NAME = osName;
    }

    protected static String searchExistingfile(final List<String> pathList) {
        for (String path : pathList) {
            if (new File(path).exists()) {
                return path;
            }
        }
        return "";
    }

}
