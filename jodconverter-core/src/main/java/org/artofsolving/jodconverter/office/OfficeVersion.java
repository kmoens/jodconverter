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
package org.artofsolving.jodconverter.office;

/**
 * Represents the OpenOffice/LibreOffice version info.
 *
 * @author Kenny Moens <kenny.moens@cipal.be>
 * @since 3.01.2.00
 */
public class OfficeVersion {
    private static final String UNKNOWN = "???";

    private String productName;
    private String version;
    private boolean useGnuStyleLongOptions = false;

    /* package */ OfficeVersion() {
        productName = UNKNOWN;
        version = UNKNOWN;
    }

    public boolean isKnown() {
        return !UNKNOWN.equals(productName);
    }

    public String getProductName() {
        return productName;
    }

    public String getVersion() {
        return version;
    }

    public int getMajorVersion() {
        String[] versionParts = version.split("\\.");
        return Integer.parseInt(versionParts[0]);
    }

    public int getMinorVersion() {
        String[] versionParts = version.split("\\.");
        if (versionParts.length > 1) {
            return Integer.parseInt(versionParts[1]);
        } else {
            return -1;
        }
    }

    public boolean useGnuStyleLongOptions() {
        return useGnuStyleLongOptions;
    }

    void setProductName(final String productName) {
        this.productName = productName;
    }

    void setVersion(final String version) {
        this.version = version;
    }

    void setUseGnuStyleLongOptions(final boolean useGnuStyleLongOptions) {
        this.useGnuStyleLongOptions = useGnuStyleLongOptions;
    }

    @Override
    public String toString() {
        return String.format("Product: %s - Version: %s - useGnuStyleLongOptions: %s", getProductName(), getVersion(), useGnuStyleLongOptions());
    }
}
