//
// JODConverter - Java OpenDocument Converter
// Copyright 2004-2012 Mirko Nasato and contributors
//
// JODConverter is Open Source software, you can redistribute it and/or
// modify it under either (at your option) of the following licenses
//
// 1. The GNU Lesser General Public License v3 (or later)
//    -> http://www.gnu.org/licenses/lgpl-3.0.txt
// 2. The Apache License, Version 2.0
//    -> http://www.apache.org/licenses/LICENSE-2.0.txt
//
package org.artofsolving.jodconverter.office;

import static org.artofsolving.jodconverter.process.ProcessManager.PID_NOT_FOUND;
import static org.artofsolving.jodconverter.process.ProcessManager.PID_UNKNOWN;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.logging.Logger;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
import org.artofsolving.jodconverter.process.ProcessManager;
import org.artofsolving.jodconverter.process.ProcessQuery;
import org.artofsolving.jodconverter.util.PlatformUtils;

class OfficeProcess {
    private static final int MAX_LENGTH = 159;

    private final File officeHome;
    private final UnoUrl unoUrl;
    private final String[] runAsArgs;
    private final File templateProfileDir;
    private final File instanceProfileDir;
    private final String instanceProfileUrl;
    private final ProcessManager processManager;
    private OfficeVersionDescriptor versionDescriptor;

    private Process process;
    private long pid = PID_UNKNOWN;
	private String commandArgPrefix;

    private final Logger logger = Logger.getLogger(getClass().getName());

    public OfficeProcess(final File officeHome, final UnoUrl unoUrl, final String[] runAsArgs, final File templateProfileDir, final File instanceProfileDir,
            final ProcessManager processManager, final boolean useGnuStyleLongOptions) {
        this.officeHome = officeHome;
        this.unoUrl = unoUrl;
        this.runAsArgs = runAsArgs;
        this.templateProfileDir = templateProfileDir;
        this.instanceProfileDir = instanceProfileDir;
        this.processManager = processManager;

        instanceProfileUrl = OfficeUtils.toUrl(instanceProfileDir);
        if (PlatformUtils.isWindows() && instanceProfileUrl.length() >= MAX_LENGTH) {
            logger.severe("The instance profile directory (" + instanceProfileUrl + ") is too long (>= " + MAX_LENGTH + " characters).");
            throw new IllegalStateException("The instance profile directory (" + instanceProfileUrl + ") is too long (>= " + MAX_LENGTH + " characters).");
        }

        if (useGnuStyleLongOptions) {
            commandArgPrefix = "--";
        } else {
            commandArgPrefix = "-";
        }
    }

    private OfficeVersionDescriptor determineOfficeVersion() {
        try {
            if (versionDescriptor != null) {
                return versionDescriptor;
            }

            File executable = OfficeUtils.getOfficeExecutable(officeHome);
            if (PlatformUtils.isWindows()) {
                versionDescriptor = OfficeVersionDescriptor.parseFromExecutableLocation(executable.getPath());
                return versionDescriptor;
            }

            List<String> command = new ArrayList<String>();
            command.add(executable.getAbsolutePath());
            command.add("-help");
            command.add("-headless");
            command.add("-nocrashreport");
            command.add("-nofirststartwizard");
            command.add("-nolockcheck");
            command.add("-nologo");
            command.add("-norestore");
            command.add("-env:UserInstallation=" + OfficeUtils.toUrl(instanceProfileDir));

            ProcessBuilder processBuilder = new ProcessBuilder(command);
            processBuilder.redirectErrorStream(true);

            Process checkProcess = processBuilder.start();
            try {
                checkProcess.waitFor();
            } catch (InterruptedException e) {
                // NOP
            }
            String versionCheckOutput = IOUtils.toString(checkProcess.getInputStream());

            versionDescriptor = OfficeVersionDescriptor.parseFromHelpOutput(versionCheckOutput);
            return versionDescriptor;
        } catch (IOException e) {
            logger.severe("Unable to determine Office version: " + e.getMessage());
            versionDescriptor =  null;
            return versionDescriptor;
        }
    }

    public void start() throws IOException {
        OfficeVersionDescriptor version = determineOfficeVersion();

        if (version != null) {
            if (version.useGnuStyleLongOptions()) {
                commandArgPrefix = "--";
            } else {
                commandArgPrefix = "-";
            }
        }
        logger.fine("OfficeProcess info:" + version.toString());
        doStart(false);
    }

    public void doStart(final boolean restart) throws IOException {
        ProcessQuery processQuery = new ProcessQuery("soffice.bin", unoUrl.getAcceptString());
        long existingPid = processManager.findPid(processQuery);
    	if (!(existingPid == PID_NOT_FOUND || existingPid == PID_UNKNOWN)) {
			throw new IllegalStateException(String.format("a process with acceptString '%s' is already running; pid %d",
			        unoUrl.getAcceptString(), existingPid));
        }
    	if (!restart) {
    	    prepareInstanceProfileDir();
    	}
        List<String> command = new ArrayList<String>();
        File executable = OfficeUtils.getOfficeExecutable(officeHome);
        if (runAsArgs != null) {
        	command.addAll(Arrays.asList(runAsArgs));
        }
        command.add(executable.getAbsolutePath());
        command.add(commandArgPrefix + "accept=" + unoUrl.getAcceptString() + ";urp;");
        command.add(commandArgPrefix + "env:UserInstallation=" + instanceProfileUrl);
        command.add(commandArgPrefix + "headless");
        command.add(commandArgPrefix + "nocrashreport");
        command.add(commandArgPrefix + "nodefault");
        command.add(commandArgPrefix + "nofirststartwizard");
        command.add(commandArgPrefix + "nolockcheck");
        command.add(commandArgPrefix + "nologo");
        command.add(commandArgPrefix + "norestore");

        ProcessBuilder processBuilder = new ProcessBuilder(command);
        processBuilder.redirectErrorStream(true);

        if (PlatformUtils.isWindows()) {
            addBasisAndUrePaths(processBuilder);
        }
        logger.info(String.format("starting process with acceptString '%s' and profileDir '%s'", unoUrl, instanceProfileDir));
        process = processBuilder.start();

		manageProcessOutputs(process);

        pid = processManager.findPid(processQuery);
        if (pid == PID_NOT_FOUND) {
            throw new IllegalStateException(String.format("process with acceptString '%s' started but its pid could not be found",
                    unoUrl.getAcceptString()));
        }
        logger.info("started process" + (pid != PID_UNKNOWN ? "; pid = " + pid : ""));
    }

    protected void manageProcessOutputs(final Process process) {
        InputStream processOut = process.getInputStream();
        InputStream processError = process.getErrorStream();

        Thread to = new Thread(new OfficeProcessStreamGobbler(processOut), "OfficeProcessStdOutThread");
        to.setDaemon(true);
        to.start();

        Thread te = new Thread(new OfficeProcessStreamGobbler(processError), "OfficeProcessStdErrThread");
        te.setDaemon(true);
        te.start();
    }

    private void prepareInstanceProfileDir() throws OfficeException {
        if (instanceProfileDir.exists()) {
            logger.warning(String.format("profile dir '%s' already exists; deleting", instanceProfileDir));
            deleteProfileDir();
        }
        if (templateProfileDir != null) {
            try {
                FileUtils.copyDirectory(templateProfileDir, instanceProfileDir);
            } catch (IOException ioException) {
                throw new OfficeException("failed to create profileDir", ioException);
            }
        }
    }

    public void deleteProfileDir() {
        if (instanceProfileDir != null) {
            try {
                FileUtils.deleteDirectory(instanceProfileDir);
            } catch (IOException ioException) {
                File oldProfileDir = new File(instanceProfileDir.getParentFile(), instanceProfileDir.getName() + ".old." + System.currentTimeMillis());
                if (instanceProfileDir.renameTo(oldProfileDir)) {
                    logger.warning("could not delete profileDir: " + ioException.getMessage() + "; renamed it to " + oldProfileDir);
                } else {
                    logger.severe("could not delete profileDir: " + ioException.getMessage());
                }
            }
        }
    }

    private void addBasisAndUrePaths(final ProcessBuilder processBuilder) throws IOException {
        // see http://wiki.services.openoffice.org/wiki/ODF_Toolkit/Efforts/Three-Layer_OOo
        File basisLink = new File(officeHome, "basis-link");
        if (!basisLink.isFile()) {
            logger.fine("no %OFFICE_HOME%/basis-link found; assuming it's OOo 2.x and we don't need to append URE and Basic paths");
            return;
        }
        String basisLinkText = FileUtils.readFileToString(basisLink).trim();
        File basisHome = new File(officeHome, basisLinkText);
        File basisProgram = new File(basisHome, "program");
        File ureLink = new File(basisHome, "ure-link");
        String ureLinkText = FileUtils.readFileToString(ureLink).trim();
        File ureHome = new File(basisHome, ureLinkText);
        File ureBin = new File(ureHome, "bin");
        Map<String,String> environment = processBuilder.environment();
        // Windows environment variables are case insensitive but Java maps are not :-/
        // so let's make sure we modify the existing key
        String pathKey = "PATH";
        for (String key : environment.keySet()) {
            if ("PATH".equalsIgnoreCase(key)) {
                pathKey = key;
            }
        }
        String path = environment.get(pathKey) + ";" + ureBin.getAbsolutePath() + ";" + basisProgram.getAbsolutePath();
        logger.fine(String.format("setting %s to \"%s\"", pathKey, path));
        environment.put(pathKey, path);
    }

    public boolean isRunning() {
        if (process == null) {
            return false;
        }
        return getExitCode() == null;
    }

    private class ExitCodeRetryable extends Retryable {
        private int exitCode;

        @Override
        protected void attempt() throws TemporaryException, Exception {
            try {
                exitCode = process.exitValue();
            } catch (IllegalThreadStateException illegalThreadStateException) {
                throw new TemporaryException(illegalThreadStateException);
            }
        }

        public int getExitCode() {
            return exitCode;
        }

    }

    public Integer getExitCode() {
        try {
            return process.exitValue();
        } catch (IllegalThreadStateException exception) {
            return null;
        }
    }

    public int getExitCode(final long retryInterval, final long retryTimeout) throws RetryTimeoutException {
        try {
            ExitCodeRetryable retryable = new ExitCodeRetryable();
            retryable.execute(retryInterval, retryTimeout);
            return retryable.getExitCode();
        } catch (RetryTimeoutException retryTimeoutException) {
            throw retryTimeoutException;
        } catch (Exception exception) {
            throw new OfficeException("could not get process exit code", exception);
        }
    }

    public int forciblyTerminate(final long retryInterval, final long retryTimeout) throws IOException, RetryTimeoutException {
        logger.info(String.format("trying to forcibly terminate process: '" + unoUrl + "'" + (pid != PID_UNKNOWN ? " (pid " + pid  + ")" : "")));
        processManager.kill(process, pid);
        return getExitCode(retryInterval, retryTimeout);
    }

    public OfficeVersionDescriptor getVersion() {
        return determineOfficeVersion();
    }
}
