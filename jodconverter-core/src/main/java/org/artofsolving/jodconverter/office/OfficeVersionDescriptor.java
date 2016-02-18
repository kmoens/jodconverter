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

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class OfficeVersionDescriptor {
    private final Logger logger = LoggerFactory.getLogger(getClass().getName());

    private String productName;
    private String version;
    private boolean useGnuStyleLongOptions = false;

    private OfficeVersionDescriptor() {
        productName = "???";
        version = "???";
    }

    public static OfficeVersionDescriptor parseFromHelpOutput(final String checkString) {
        OfficeVersionDescriptor desc = new OfficeVersionDescriptor();

        String productLine = null;
        String[] lines = checkString.split("\\n");
        for (String line : lines) {
            if (line.contains("--help")) {
                desc.useGnuStyleLongOptions = true;
            }
            String lowerLine = line.trim().toLowerCase();
            if (lowerLine.startsWith("openoffice") || lowerLine.startsWith("libreoffice")) {
                productLine = line.trim();
            }
        }
        if (productLine != null) {
            String[] parts = productLine.split(" ");
            if (parts.length > 0) {
                desc.productName = parts[0];
            } else {
                desc.productName = "???";
            }
            if (parts.length > 1) {
                desc.version = parts[1];
            } else {
                desc.version = "???";
            }
        } else {
            desc.productName = "???";
            desc.version = "???";
        }

        return desc;
    }

    public static OfficeVersionDescriptor parseFromExecutableLocation(final String path) {
        OfficeVersionDescriptor desc = new OfficeVersionDescriptor();

        if (path.toLowerCase().contains("openoffice")) {
            desc.productName = "OpenOffice";
            desc.useGnuStyleLongOptions = false;
        }
        if (path.toLowerCase().contains("libreoffice")) {
            desc.productName = "LibreOffice";
            desc.useGnuStyleLongOptions = true;
        }

        String[] versionsToCheck = { "5.1", "5", "4.1.2", "4.1.1", "4.1", "4", "3.9", "3.8", "3.7", "3.6", "3.5", "3.4", "3.3", "3.2", "3.1", "3" };

        for (String v : versionsToCheck) {
            if (path.contains(v)) {
                desc.version = v;
                break;
            }
        }

        return desc;
    }

    public String getProductName() {
        return productName;
    }

    public void setProductName(final String productName) {
        this.productName = productName;
    }

    public String getVersion() {
        return version;
    }

    public void setVersion(final String version) {
        this.version = version;
    }

    public boolean useGnuStyleLongOptions() {
        return useGnuStyleLongOptions;
    }

    @Override
    public String toString() {
        return String.format("Product: %s - Version: %s - useGnuStyleLongOptions: %s", getProductName(), getVersion(), useGnuStyleLongOptions());
    }
}
