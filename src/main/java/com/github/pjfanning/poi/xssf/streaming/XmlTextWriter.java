package com.github.pjfanning.poi.xssf.streaming;

import java.io.IOException;
import java.io.Writer;

/**
 * Writes XML text content directly, without building an XMLBeans object to serialize.
 * <p>
 * The output is byte for byte what {@code XSSFRichTextString} plus {@code CTRst.xmlText}
 * produces for a plain (unformatted) string, so switching to it does not change the files
 * that get written.
 * </p>
 */
class XmlTextWriter {

    private XmlTextWriter() {}

    /**
     * Writes a {@code <t>} element holding the supplied text.
     *
     * @param writer the writer to write to
     * @param text the text to write
     * @throws IOException if the writer throws
     */
    static void writeTElement(Writer writer, String text) throws IOException {
        if (text.isEmpty()) {
            writer.write("<t/>");
            return;
        }
        writer.write(needsSpacePreserve(text) ? "<t xml:space=\"preserve\">" : "<t>");
        writeEscaped(writer, text);
        writer.write("</t>");
    }

    /**
     * POI marks the element with {@code xml:space="preserve"} when the text starts or ends with
     * whitespace, so that the leading/trailing whitespace survives a round trip. It uses
     * {@link Character#isWhitespace}, which is broader than the XML whitespace set - it also
     * covers vertical tab and the file/group/record/unit separators - and it is applied to the
     * original text, before any character that XML cannot represent has been replaced.
     */
    static boolean needsSpacePreserve(String text) {
        return !text.isEmpty()
                && (Character.isWhitespace(text.charAt(0))
                    || Character.isWhitespace(text.charAt(text.length() - 1)));
    }

    /**
     * Writes the text as XML character data, escaping what has to be escaped.
     *
     * @param writer the writer to write to
     * @param text the text to escape
     * @throws IOException if the writer throws
     */
    static void writeEscaped(Writer writer, String text) throws IOException {
        final int len = text.length();
        int start = 0;
        for (int i = 0; i < len; i++) {
            final char c = text.charAt(i);
            final String replacement;
            switch (c) {
                case '&':
                    replacement = "&amp;";
                    break;
                case '<':
                    replacement = "&lt;";
                    break;
                case '\r':
                    // a literal CR would be normalised to LF when the file is read back
                    replacement = "&#13;";
                    break;
                case '>':
                    // only needs escaping where it would otherwise close a CDATA section
                    if (i >= 2 && text.charAt(i - 1) == ']' && text.charAt(i - 2) == ']') {
                        replacement = "&gt;";
                    } else {
                        continue;
                    }
                    break;
                default:
                    if (isValidXmlChar(c)) {
                        continue;
                    }
                    // matches how XMLBeans handles characters that XML 1.0 cannot represent
                    replacement = "?";
                    break;
            }
            if (i > start) {
                writer.write(text, start, i - start);
            }
            writer.write(replacement);
            start = i + 1;
        }
        if (start < len) {
            writer.write(text, start, len - start);
        }
    }

    private static boolean isValidXmlChar(char c) {
        if (c >= 0x20) {
            return c != 0xFFFE && c != 0xFFFF;
        }
        return c == '\t' || c == '\n' || c == '\r';
    }
}
