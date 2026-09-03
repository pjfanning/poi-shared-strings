package com.github.pjfanning.poi.xssf.streaming;

import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.xmlbeans.XmlOptions;
import org.junit.Test;

import javax.xml.namespace.QName;
import java.io.StringWriter;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Random;

import static org.apache.poi.xssf.usermodel.XSSFRelation.NS_SPREADSHEETML;
import static org.junit.Assert.assertEquals;

/**
 * XmlTextWriter replaces an XSSFRichTextString round trip, so it has to produce exactly what
 * XMLBeans produced. These tests compare the two outputs directly.
 */
public class TestXmlTextWriter {

    private static final XmlOptions SI_SAVE_OPTIONS = new XmlOptions();
    static {
        SI_SAVE_OPTIONS.setCharacterEncoding("UTF-8");
        SI_SAVE_OPTIONS.setSaveAggressiveNamespaces();
        SI_SAVE_OPTIONS.setUseDefaultNamespace(true);
        SI_SAVE_OPTIONS.setSaveImplicitNamespaces(Collections.singletonMap("", NS_SPREADSHEETML));
        SI_SAVE_OPTIONS.setSaveSyntheticDocumentElement(new QName(NS_SPREADSHEETML, "si"));
    }

    private static final XmlOptions TEXT_SAVE_OPTIONS = new XmlOptions();
    static {
        TEXT_SAVE_OPTIONS.setCharacterEncoding("UTF-8");
        TEXT_SAVE_OPTIONS.setSaveAggressiveNamespaces();
        TEXT_SAVE_OPTIONS.setUseDefaultNamespace(true);
        TEXT_SAVE_OPTIONS.setSaveImplicitNamespaces(Collections.singletonMap("", NS_SPREADSHEETML));
        TEXT_SAVE_OPTIONS.setSaveSyntheticDocumentElement(new QName(NS_SPREADSHEETML, "text"));
    }

    private static String viaXmlBeans(String text) {
        return new XSSFRichTextString(text).getCTRst().xmlText(SI_SAVE_OPTIONS);
    }

    private static String viaXmlTextWriter(String text) throws Exception {
        StringWriter sw = new StringWriter();
        sw.write("<si>");
        XmlTextWriter.writeTElement(sw, text);
        sw.write("</si>");
        return sw.toString();
    }

    private static String withChar(String prefix, int code, String suffix) {
        return prefix + (char) code + suffix;
    }

    private static void assertMatchesXmlBeans(String text) throws Exception {
        assertEquals("output for " + describe(text), viaXmlBeans(text), viaXmlTextWriter(text));
    }

    private static String describe(String s) {
        StringBuilder b = new StringBuilder();
        for (char c : s.toCharArray()) {
            if (c < 0x20 || c > 0x7e) {
                b.append(String.format("\\u%04x", (int) c));
            } else {
                b.append(c);
            }
        }
        return "\"" + b + "\"";
    }

    @Test
    public void testPlainText() throws Exception {
        assertMatchesXmlBeans("plain");
        assertMatchesXmlBeans("");
        assertMatchesXmlBeans("a string with spaces inside");
    }

    @Test
    public void testLeadingAndTrailingWhitespaceIsPreserved() throws Exception {
        assertMatchesXmlBeans(" lead");
        assertMatchesXmlBeans("trail ");
        assertMatchesXmlBeans("  both  ");
        assertMatchesXmlBeans(" ");
        assertMatchesXmlBeans("\tleadtab");
        assertMatchesXmlBeans("trailtab\t");
        assertMatchesXmlBeans("\nleadnl");
        assertMatchesXmlBeans("trailnl\n");
        assertMatchesXmlBeans("\rleadcr");
        assertMatchesXmlBeans("trailcr\r");
    }

    @Test
    public void testMarkupCharacters() throws Exception {
        assertMatchesXmlBeans("amp&");
        assertMatchesXmlBeans("lt<");
        assertMatchesXmlBeans("gt>");
        assertMatchesXmlBeans("quote\"");
        assertMatchesXmlBeans("apos'");
        assertMatchesXmlBeans("<tag attr=\"v\">body</tag>");
        assertMatchesXmlBeans("&amp; already escaped");
        assertMatchesXmlBeans("all &<>\"' of them");
    }

    @Test
    public void testCdataClose() throws Exception {
        assertMatchesXmlBeans("a]]>b");
        assertMatchesXmlBeans("a]>b");
        assertMatchesXmlBeans("a>>b");
        assertMatchesXmlBeans("]]>");
        assertMatchesXmlBeans("]]]>");
        assertMatchesXmlBeans("]]>]]>");
    }

    @Test
    public void testControlAndNonXmlCharacters() throws Exception {
        assertMatchesXmlBeans(withChar("nul", 0x00, "x"));
        assertMatchesXmlBeans(withChar("ctl", 0x01, "x"));
        assertMatchesXmlBeans(withChar("vtab", 0x0b, "x"));
        assertMatchesXmlBeans(withChar("ff", 0x0c, "x"));
        assertMatchesXmlBeans(withChar("esc", 0x1b, "x"));
        assertMatchesXmlBeans(withChar("us", 0x1f, "x"));
        assertMatchesXmlBeans(withChar("del", 0x7f, "x"));
        assertMatchesXmlBeans(withChar("nel", 0x85, "x"));
        assertMatchesXmlBeans(withChar("nonchar", 0xfffe, "x"));
        assertMatchesXmlBeans(withChar("nonchar", 0xffff, "x"));
        assertMatchesXmlBeans(withChar("repl", 0xfffd, "x"));
    }

    @Test
    public void testUnicode() throws Exception {
        assertMatchesXmlBeans("\u00e9accent");
        assertMatchesXmlBeans("\u65e5\u672c\u8a9e");
        assertMatchesXmlBeans("emoji\ud83d\ude00here");
        assertMatchesXmlBeans("lone\ud800surrogate");
    }

    @Test
    public void testCommentTextElementMatchesXmlBeans() throws Exception {
        // the comments table writes the same text wrapped in <text> rather than <si>
        final String[] cases = {
            "plain", "", " lead", "trail ", "amp & here", "lt < gt >", "quote \" apos '",
            "cdata ]]> close", "\u00e9\u65e5", "multi\nline", "carriage\rreturn"
        };
        for (String text : cases) {
            final String expected =
                    new XSSFRichTextString(text).getCTRst().xmlText(TEXT_SAVE_OPTIONS);
            final StringWriter sw = new StringWriter();
            sw.write("<text>");
            XmlTextWriter.writeTElement(sw, text);
            sw.write("</text>");
            assertEquals("comment text for " + describe(text), expected, sw.toString());
        }
    }

    @Test
    public void testRandomStringsMatchXmlBeans() throws Exception {
        // fuzz over an alphabet loaded with the characters that need special handling
        final char[] alphabet = {
            'a', 'b', 'Z', '0', ' ', ' ', '\t', '\n', '\r', '&', '<', '>', ']', '"', '\'',
            ' ', (char) 0x00, (char) 0x01, (char) 0x0b, (char) 0x1b, (char) 0x7f,
            (char) 0xe9, (char) 0x65e5, (char) 0xfffe, (char) 0xfffd
        };
        final Random random = new Random(20260903L);
        final List<String> failures = new ArrayList<>();
        for (int i = 0; i < 5000; i++) {
            final StringBuilder sb = new StringBuilder();
            final int len = random.nextInt(12);
            for (int j = 0; j < len; j++) {
                sb.append(alphabet[random.nextInt(alphabet.length)]);
            }
            final String text = sb.toString();
            final String expected = viaXmlBeans(text);
            final String actual = viaXmlTextWriter(text);
            if (!expected.equals(actual)) {
                failures.add(describe(text) + " expected " + expected + " but got " + actual);
            }
        }
        assertEquals("mismatches: " + failures, 0, failures.size());
    }
}
