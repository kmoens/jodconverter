package org.artofsolving.jodconverter.office;

import static org.assertj.core.api.Assertions.assertThat;

import java.io.File;

import org.artofsolving.jodconverter.process.PureJavaProcessManager;
import org.artofsolving.jodconverter.util.PlatformUtils;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.ExpectedException;
import org.powermock.reflect.Whitebox;

public class OfficeProcessTest {
    private static final String PATH_TOO_LONG = "c:/cipal_java/Apache Software Foundation/tomcat6/temp/OLYMPUS40b462c1-ed8a-479f-904d-aa8e4822f410/.jodconverter_socket_host-127.0.0.1_port-7201abcd";
    private static final String PATH_BORDER = "c:/cipal_java/Apache Software Foundation/tomcat6/temp/OLYMPUS40b462c1-ed8a-479f-904d-aa8e4822f410/.jodconverter_socket_host-127.0.0.1_port-7201abc";

    @Rule
    public ExpectedException exception = ExpectedException.none();

    @Test
    public void testInstanceProfileDir() {
        // when
        OfficeProcess op = new OfficeProcess(new File("/foo"), UnoUrl.socket(1234), new String[0], new File("/templates"), new File(PATH_BORDER), new PureJavaProcessManager(), false);

        // then
        assertThat(Whitebox.getInternalState(op, "instanceProfileDir")).isInstanceOf(File.class);
        assertThat((File) Whitebox.getInternalState(op, "instanceProfileDir")).isEqualTo(new File(PATH_BORDER));
    }

    @Test
    public void testInstanceProfileDirThrowsExceptionIfTooLongOnWindows() {
        // exception
        exception.expect(IllegalStateException.class);
        exception.expectMessage("The instance profile directory (" + OfficeUtils.toUrl(new File(PATH_TOO_LONG)) + ") is too long (>= 159 characters).");

        // given
        PlatformUtils.set("windows");

        // when
        OfficeProcess op = new OfficeProcess(new File("/foo"), UnoUrl.socket(1234), new String[0], new File("/templates"), new File(PATH_TOO_LONG), new PureJavaProcessManager(), false);
    }

    @Test
    public void testInstanceProfileDirThrowsExceptionIfNotTooLongOnLinux() {
        // given
        PlatformUtils.set("linux");

        // when
        OfficeProcess op = new OfficeProcess(new File("/foo"), UnoUrl.socket(1234), new String[0], new File("/templates"), new File(PATH_TOO_LONG), new PureJavaProcessManager(), false);

        // then
        assertThat(Whitebox.getInternalState(op, "instanceProfileDir")).isInstanceOf(File.class);
        assertThat((File) Whitebox.getInternalState(op, "instanceProfileDir")).isEqualTo(new File(PATH_TOO_LONG));
    }

}
