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

import java.io.File;
import java.util.concurrent.ArrayBlockingQueue;
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.TimeUnit;

import org.artofsolving.jodconverter.process.ProcessManager;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

class ProcessPoolOfficeManager implements OfficeManager {
    private final BlockingQueue<PooledOfficeManager> pool;
    private final PooledOfficeManager[] pooledManagers;
    private final OfficeVersion officeVersion;
    private final long taskQueueTimeout;

    private volatile boolean running = false;

    private final Logger logger = LoggerFactory.getLogger(ProcessPoolOfficeManager.class.getName());

    public ProcessPoolOfficeManager(final File officeHome, final UnoUrl[] unoUrls, final String[] runAsArgs, final File templateProfileDir, final File workDir,
            final long retryTimeout, final long taskQueueTimeout, final long taskExecutionTimeout, final int maxTasksPerProcess,
            final ProcessManager processManager, final boolean useGnuStyleLongOptions) {
		this.taskQueueTimeout = taskQueueTimeout;
        pool = new ArrayBlockingQueue<PooledOfficeManager>(unoUrls.length);
        pooledManagers = new PooledOfficeManager[unoUrls.length];
        for (int i = 0; i < unoUrls.length; i++) {
            PooledOfficeManagerSettings settings = new PooledOfficeManagerSettings(unoUrls[i]);
            settings.setRunAsArgs(runAsArgs);
            settings.setTemplateProfileDir(templateProfileDir);
            settings.setWorkDir(workDir);
            settings.setOfficeHome(officeHome);
            settings.setRetryTimeout(retryTimeout);
            settings.setTaskExecutionTimeout(taskExecutionTimeout);
            settings.setMaxTasksPerProcess(maxTasksPerProcess);
            settings.setProcessManager(processManager);
            settings.setUseGnuStyleLongOptions(useGnuStyleLongOptions);
            pooledManagers[i] = new PooledOfficeManager(settings);
        }

        OfficeVersionDetector versionDetector = new OfficeVersionDetector(officeHome);
        officeVersion = versionDetector.getVersion();
        logger.info("ProcessManager implementation is " + processManager.getClass().getSimpleName());
    }

    @Override
    public synchronized void start() throws OfficeException {
        for (PooledOfficeManager pooledManager : pooledManagers) {
            pooledManager.start();
            releaseManager(pooledManager);
        }
        running = true;
    }

    @Override
    public void execute(final OfficeTask task) throws IllegalStateException, OfficeException {
        if (!running) {
            throw new IllegalStateException("this OfficeManager is currently stopped");
        }
        PooledOfficeManager manager = null;
        try {
            manager = acquireManager();
            if (manager == null) {
                throw new OfficeException("no office manager available");
            }
            manager.execute(task);
        } finally {
            if (manager != null) {
                releaseManager(manager);
            }
        }
    }

    @Override
    public synchronized void stop() throws OfficeException {
        running = false;
        logger.info("stopping");
        pool.clear();
        for (PooledOfficeManager pooledManager : pooledManagers) {
            pooledManager.stop();
        }
        logger.info("stopped");
    }

    private PooledOfficeManager acquireManager() {
        try {
            return pool.poll(taskQueueTimeout, TimeUnit.MILLISECONDS);
        } catch (InterruptedException interruptedException) {
            throw new OfficeException("interrupted", interruptedException);
        }
    }

    private void releaseManager(final PooledOfficeManager manager) {
        try {
            pool.put(manager);
        } catch (InterruptedException interruptedException) {
            throw new OfficeException("interrupted", interruptedException);
        }
    }

	@Override
    public boolean isRunning() {
		return running;
	}

    @Override
    public OfficeVersion getVersion() {
        return officeVersion;
    }
}
