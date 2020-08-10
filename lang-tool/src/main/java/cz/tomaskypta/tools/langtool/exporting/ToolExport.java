package cz.tomaskypta.tools.langtool.exporting;

import java.io.*;
import java.util.*;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.util.CellRangeAddress;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;


public class ToolExport {

    private static final String EXCEL_EXTENSION = ".xls";

    private static final String DIR_VALUES = "values";
    private static final String[] POTENTIAL_RES_DIRS = new String[]{"res", "src/main/res"};

    private DocumentBuilder builder;
    private File outExcelFile;
    private String project;
    private Map<String, Map<String, Integer>> fileKeys;
    private Map<String, Boolean> untranslatableMap;
    private PrintStream out;
    private ExportConfig mConfig;
    private Set<String> sAllowedFiles = new HashSet<String>();

    {
        sAllowedFiles.add("strings.xml");
    }

    public ToolExport(PrintStream out) throws ParserConfigurationException {
        DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
        builder = dbf.newDocumentBuilder();
        this.untranslatableMap = new HashMap<String, Boolean>();
        this.out = out == null ? System.out : out;
    }

    public static void run(ExportConfig config) throws SAXException,
            IOException, ParserConfigurationException {
        run(null, config);
    }

    public static void run(PrintStream out, ExportConfig config) throws SAXException, IOException, ParserConfigurationException {
        ToolExport tool = new ToolExport(out);
        if (StringUtils.isEmpty(config.inputExportProject)) {
            tool.out.println("Cannot export, missing config");
            return;
        }
        File project = new File(config.inputExportProject);
        if (StringUtils.isEmpty(config.outputFile)) {
            config.outputFile = "exported_strings_" + System.currentTimeMillis() + EXCEL_EXTENSION;
        }
        else if (!StringUtils.endsWith(config.outputFile, EXCEL_EXTENSION)) {
            config.outputFile += EXCEL_EXTENSION;
        }
        tool.outExcelFile = new File(config.outputFile);

        if (!tool.outExcelFile.exists()) {
            try {
                tool.outExcelFile.createNewFile();
            } catch (Exception e) {
                System.err.println("Cannot create file in position: " + tool.outExcelFile.getAbsolutePath());
                e.printStackTrace();
            }
        }
        tool.outExcelFile.createNewFile();

        tool.project = project.getName();
        tool.mConfig = config;
        tool.sAllowedFiles.addAll(config.additionalResources);
        tool.export(project);
    }

    private void export(File project) throws SAXException, IOException {
        File res = findResourceDir(project);
        if (res == null) {
            System.err.println("Cannot find resource directory.");
            return;
        }
        this.fileKeys = new HashMap<String, Map<String, Integer>>();
        for (File dir : res.listFiles()) {
            if (!dir.isDirectory() || !dir.getName().startsWith(DIR_VALUES)) {
                continue;
            }
            if (dir.getName().equals(DIR_VALUES)) {
                exportDefLang(dir);
                break;
            }
        }
        for (File dir : res.listFiles()) {
            if (!dir.isDirectory() || !dir.getName().startsWith(DIR_VALUES)) {
                continue;
            }
            String dirName = dir.getName();
            if (!dirName.equals(DIR_VALUES))  {
                int index = dirName.indexOf('-');
                if (index == -1)
                    continue;
                String lang = dirName.substring(index + 1);
                exportLang(lang, dir);
            }
        }
    }

    private File findResourceDir(File project) {
        List<File> availableResDirs = new LinkedList<File>();
        for (String potentialResDir : POTENTIAL_RES_DIRS) {
            File res = new File(project, potentialResDir);
            if (res.exists()) {
                availableResDirs.add(res);
            }
        }
        if (!availableResDirs.isEmpty()) {
            return availableResDirs.get(0);
        }
        return null;
    }

    private void exportLang(String lang, File valueDir) throws IOException, SAXException {
        for (String fileName : sAllowedFiles) {
            File stringFile = new File(valueDir, fileName);
            if (!stringFile.exists()) {
                continue;
            }
            Map<String, Integer> keysIndex = this.fileKeys.get(fileName);
            exportLangToExcel(project, lang, stringFile, getStrings(stringFile), outExcelFile, keysIndex);
        }
    }

    private void exportDefLang(File valueDir) throws IOException, SAXException {
        HSSFWorkbook wb = new HSSFWorkbook();
        for (String fileName : sAllowedFiles) {
            Map<String, Integer> keys = new HashMap<String, Integer>();
            FileOutputStream outFile = new FileOutputStream(outExcelFile);
            HSSFSheet sheet;
            sheet = wb.createSheet(fileName);
            sheet.createRow(0);
            createTilte(wb, sheet);
            addLang2Tilte(wb, sheet, "default");
            addTranslatable(wb, sheet);
            sheet.createFreezePane(1, 1);
            wb.write(outFile);
            outFile.close();
            if (this.fileKeys.get(fileName) == null) {
                this.fileKeys.put(fileName, keys);
            }
        }
        for (String fileName : sAllowedFiles) {
            Map<String, Integer> keys = this.fileKeys.get(fileName);
            File stringFile = new File(valueDir, fileName);
            if (!stringFile.exists()) {
                continue;
            }
            keys.putAll(exportDefLangToExcel(1, project, stringFile, getStrings(stringFile), outExcelFile));
        }
    }

    private NodeList getStrings(File f) throws SAXException, IOException {
        Document dom = builder.parse(f);
        return dom.getDocumentElement().getChildNodes();
    }

    private static HSSFCellStyle createTilteStyle(HSSFWorkbook wb) {
        HSSFFont bold = wb.createFont();
        bold.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);

        HSSFCellStyle style = wb.createCellStyle();
        style.setFont(bold);
        style.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        style.setWrapText(true);

        return style;
    }

    private static HSSFCellStyle createCommentStyle(HSSFWorkbook wb) {

        HSSFFont commentFont = wb.createFont();
        commentFont.setColor(HSSFColor.GREEN.index);
        commentFont.setItalic(true);
        commentFont.setFontHeightInPoints((short)12);

        HSSFCellStyle commentStyle = wb.createCellStyle();
        commentStyle.setFont(commentFont);
        return commentStyle;
    }

    private static HSSFCellStyle createSectionCommentStyle(HSSFWorkbook wb) {

        HSSFFont commentFont = wb.createFont();
        commentFont.setColor(HSSFColor.GREEN.index);
        commentFont.setItalic(true);
        commentFont.setFontHeightInPoints((short)12);

        HSSFCellStyle commentStyle = wb.createCellStyle();
        commentStyle.setFillForegroundColor(HSSFColor.GREEN.LIGHT_GREEN.index);
        commentStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        commentStyle.setFont(commentFont);
        return commentStyle;
    }

    private static HSSFCellStyle createUntranslatableStyle(HSSFWorkbook wb) {
        HSSFCellStyle textStyle = wb.createCellStyle();
        textStyle.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
        textStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        textStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        HSSFFont font = wb.createFont();
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        textStyle.setFont(font);
        return textStyle;
    }

    private static HSSFCellStyle createStringReferenceStyle(HSSFWorkbook wb) {
        HSSFCellStyle textStyle = wb.createCellStyle();
        HSSFFont font = wb.createFont();
        font.setItalic(true);
        font.setFontHeightInPoints((short)12);
        font.setColor(HSSFColor.GREY_40_PERCENT.index);
        textStyle.setFont(font);
        return textStyle;
    }

    private static HSSFCellStyle createPlurarStyle(HSSFWorkbook wb) {

        HSSFFont commentFont = wb.createFont();
        commentFont.setColor(HSSFColor.GREY_50_PERCENT.index);
        commentFont.setItalic(true);
        commentFont.setFontHeightInPoints((short)12);

        HSSFCellStyle commentStyle = wb.createCellStyle();
        commentStyle.setFont(commentFont);
        return commentStyle;
    }

    private static HSSFCellStyle createKeyStyle(HSSFWorkbook wb) {
        HSSFFont bold = wb.createFont();
        bold.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        bold.setFontHeightInPoints((short)11);

        HSSFCellStyle keyStyle = wb.createCellStyle();
        keyStyle.setFont(bold);

        return keyStyle;
    }

    private static HSSFCellStyle createTextStyle(HSSFWorkbook wb) {
        HSSFFont plain = wb.createFont();
        plain.setFontHeightInPoints((short)12);

        HSSFCellStyle textStyle = wb.createCellStyle();
        textStyle.setFont(plain);

        return textStyle;
    }

    private static HSSFCellStyle createMissedStyle(HSSFWorkbook wb) {

        HSSFCellStyle style = wb.createCellStyle();
        style.setFillForegroundColor(HSSFColor.RED.index);
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

        return style;
    }

    private static void createTilte(HSSFWorkbook wb, HSSFSheet sheet) {
        HSSFRow titleRow = sheet.getRow(0);

        HSSFCell cell = titleRow.createCell(0);
        cell.setCellStyle(createTilteStyle(wb));
        cell.setCellValue("KEY");

        sheet.setColumnWidth(cell.getColumnIndex(), (40 * 256));
    }

    private static void addTranslatable(HSSFWorkbook wb, HSSFSheet sheet) {
        HSSFRow titleRow = sheet.getRow(0);

        HSSFCell cell = titleRow.createCell(2);
        cell.setCellStyle(createTilteStyle(wb));
        cell.setCellValue("Untranslatable");

        sheet.setColumnWidth(cell.getColumnIndex(), (18 * 256));
    }

    private static void addLang2Tilte(HSSFWorkbook wb, HSSFSheet sheet, String lang) {
        HSSFRow titleRow = sheet.getRow(0);
        HSSFCell lastCell = titleRow.getCell((int)titleRow.getLastCellNum() - 1);
        if (lang.equals(lastCell.getStringCellValue())) {
            // language column already exists
            return;
        }
        HSSFCell cell = titleRow.createCell((int)titleRow.getLastCellNum());
        cell.setCellStyle(createTilteStyle(wb));
        cell.setCellValue(lang);

        sheet.setColumnWidth(cell.getColumnIndex(), (60 * 256));
    }


    private Map<String, Integer> exportDefLangToExcel(int rowIndex, String project, File src, NodeList strings, File f) throws FileNotFoundException, IOException {
        out.println();
        out.println("Start processing DEFAULT language " + src.getName());

        Map<String, Integer> keys = new HashMap<String, Integer>();

        HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(f));

        HSSFCellStyle sectionCommentStyle = createSectionCommentStyle(wb);
        HSSFCellStyle commentStyle = createCommentStyle(wb);
        HSSFCellStyle plurarStyle = createPlurarStyle(wb);
        HSSFCellStyle keyStyle = createKeyStyle(wb);
        HSSFCellStyle textStyle = createTextStyle(wb);
        HSSFCellStyle untranslatableStyle = createUntranslatableStyle(wb);
        HSSFCellStyle stringReferenceStyle = createStringReferenceStyle(wb);

        HSSFSheet sheet = wb.getSheet(src.getName());


        for (int i = 0; i < strings.getLength(); i++) {
            Node item = strings.item(i);
            if (item.getNodeType() == Node.TEXT_NODE) {

            }
            if (item.getNodeType() == Node.COMMENT_NODE) {
                String content = item.getTextContent();
                boolean sectionTitle = content.startsWith("$");
                content = content.replace("$", "");
                HSSFRow row = sheet.createRow(rowIndex++);
                HSSFCell cell = row.createCell(0);
                cell.setCellValue(String.format("/** %s **/", content));
                cell.setCellStyle(sectionTitle? sectionCommentStyle: commentStyle);
                sheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 0, 255));
            }

            if ("string".equals(item.getNodeName())) {
                Node translatableNode = item.getAttributes().getNamedItem("translatable");
                boolean untranslatable = translatableNode != null && "false".equals(translatableNode.getNodeValue());
                String key = item.getAttributes().getNamedItem("name").getNodeValue();
                this.untranslatableMap.put(key, untranslatable);
                if (mConfig.isIgnoredKey(key)) {
                    continue;
                }
                keys.put(key, rowIndex);

                HSSFRow row = sheet.createRow(rowIndex++);

                HSSFCell cell = row.createCell(0);
                cell.setCellValue(key);
                cell.setCellStyle(keyStyle);

                cell = row.createCell(1);
                boolean referenced = item.getTextContent().startsWith("@string/");
                cell.setCellStyle(referenced? stringReferenceStyle: textStyle);
                cell.setCellValue(item.getTextContent());

                cell = row.createCell(2);
                cell.setCellStyle(untranslatable? untranslatableStyle: textStyle);
                cell.setCellValue(untranslatable? "✓": "");
            } else if ("plurals".equals(item.getNodeName())) {
                String key = item.getAttributes().getNamedItem("name").getNodeValue();
                if (mConfig.isIgnoredKey(key)) {
                    continue;
                }
                String plurarName = key;

                HSSFRow row = sheet.createRow(rowIndex++);
                HSSFCell cell = row.createCell(0);
                cell.setCellValue(String.format("//plurals: %s", plurarName));
                cell.setCellStyle(plurarStyle);

                NodeList items = item.getChildNodes();
                for (int j = 0; j < items.getLength(); j++) {
                    Node plurarItem = items.item(j);
                    if ("item".equals(plurarItem.getNodeName())) {
                        String itemKey = plurarName + "#" + plurarItem.getAttributes().getNamedItem("quantity").getNodeValue();
                        boolean untranslatable = plurarItem.getTextContent().startsWith("@string/");
                        this.untranslatableMap.put(itemKey, untranslatable);

                        keys.put(itemKey, rowIndex);

                        HSSFRow itemRow = sheet.createRow(rowIndex++);

                        HSSFCell itemCell = itemRow.createCell(0);
                        itemCell.setCellValue(itemKey);
                        itemCell.setCellStyle(keyStyle);

                        itemCell = itemRow.createCell(1);
                        itemCell.setCellStyle(untranslatable? stringReferenceStyle: textStyle);
                        itemCell.setCellValue(plurarItem.getTextContent());

                        itemCell = itemRow.createCell(2);
                        itemCell.setCellStyle(untranslatable? untranslatableStyle: textStyle);
                        itemCell.setCellValue(untranslatable? "✓": "");
                    }
                }
            } else if ("string-array".equals(item.getNodeName())) {
                String key = item.getAttributes().getNamedItem("name").getNodeValue();
                if (mConfig.isIgnoredKey(key)) {
                    continue;
                }
                NodeList arrayItems = item.getChildNodes();
                for (int j = 0, k = 0; j < arrayItems.getLength(); j++) {
                    Node arrayItem = arrayItems.item(j);
                    if ("item".equals(arrayItem.getNodeName())) {
                        String itemKey = key + "[" + k++ + "]";
                        boolean untranslatable = arrayItem.getTextContent().startsWith("@string/");
                        this.untranslatableMap.put(itemKey, untranslatable);
                        keys.put(itemKey, rowIndex);

                        HSSFRow itemRow = sheet.createRow(rowIndex++);

                        HSSFCell itemCell = itemRow.createCell(0);
                        itemCell.setCellValue(itemKey);
                        itemCell.setCellStyle(keyStyle);

                        itemCell = itemRow.createCell(1);
                        itemCell.setCellStyle(untranslatable? stringReferenceStyle: textStyle);
                        itemCell.setCellValue(arrayItem.getTextContent());

                        itemCell = itemRow.createCell(2);
                        itemCell.setCellStyle(untranslatable? untranslatableStyle: textStyle);
                        itemCell.setCellValue(untranslatable? "✓": "");
                    }
                }
            }
        }

        FileOutputStream outFile = new FileOutputStream(f);
        wb.write(outFile);
        outFile.close();

        out.println("DEFAULT language was precessed");
        return keys;
    }

    private void exportLangToExcel(String project, String lang, File src, NodeList strings, File f, Map<String, Integer> keysIndex) throws FileNotFoundException, IOException {
        out.println();
        out.println(String.format("Start processing: '%s'", lang) + " " + src.getName());
        Set<String> missedKeys = new HashSet<String>(keysIndex.keySet());

        HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(f));

        HSSFCellStyle textStyle = createTextStyle(wb);

        HSSFSheet sheet = wb.getSheet(src.getName());
        addLang2Tilte(wb, sheet, lang);

        HSSFRow titleRow = sheet.getRow(0);
        int lastColumnIdx = (int)titleRow.getLastCellNum() - 1;

        for (int i = 0; i < strings.getLength(); i++) {
            Node item = strings.item(i);

            if ("string".equals(item.getNodeName())) {
                Node translatable = item.getAttributes().getNamedItem("translatable");
                if (translatable != null && "false".equals(translatable.getNodeValue())) {
                    continue;
                }
                String key = item.getAttributes().getNamedItem("name").getNodeValue();
                Integer index = keysIndex.get(key);
                if (index == null) {
                    out.println("\t" + key + " - row does not exist");
                    continue;
                }

                missedKeys.remove(key);
                HSSFRow row = sheet.getRow(index);

                HSSFCell cell = row.createCell(lastColumnIdx);
                cell.setCellValue(item.getTextContent());
                cell.setCellStyle(textStyle);
            } else if ("plurals".equals(item.getNodeName())) {
                String key = item.getAttributes().getNamedItem("name").getNodeValue();
                String plurarName = key;

                NodeList items = item.getChildNodes();
                for (int j = 0; j < items.getLength(); j++) {
                    Node plurarItem = items.item(j);
                    if ("item".equals(plurarItem.getNodeName())) {
                        key = plurarName + "#" + plurarItem.getAttributes().getNamedItem("quantity").getNodeValue();
                        Integer index = keysIndex.get(key);
                        if (index == null) {
                            out.println("\t" + key + " - row does not exist");
                            continue;
                        }
                        missedKeys.remove(key);

                        HSSFRow row = sheet.getRow(index);

                        HSSFCell cell = row.createCell(lastColumnIdx);
                        cell.setCellValue(plurarItem.getTextContent());
                        cell.setCellStyle(textStyle);
                    }
                }
            } else if ("string-array".equals(item.getNodeName())) {
                String key = item.getAttributes().getNamedItem("name").getNodeValue();
                NodeList arrayItems = item.getChildNodes();
                for (int j = 0, k = 0; j < arrayItems.getLength(); j++) {
                    Node arrayItem = arrayItems.item(j);
                    if ("item".equals(arrayItem.getNodeName())) {
                        String itemKey = key + "[" + k++ + "]";
                        Integer rowIndex = keysIndex.get(itemKey);
                        if (rowIndex == null) {
                            out.println("\t" + key + " - row does not exist");
                            continue;
                        }
                        missedKeys.remove(key);

                        HSSFRow itemRow = sheet.getRow(rowIndex);

                        HSSFCell cell = itemRow.createCell(lastColumnIdx);
                        cell.setCellValue(arrayItem.getTextContent());
                        cell.setCellStyle(textStyle);
                    }
                }
            }
        }

        HSSFCellStyle missedStyle = createMissedStyle(wb);
        HSSFCellStyle untranstableStyle = createUntranslatableStyle(wb);
        if (!missedKeys.isEmpty()) {
            out.println("  MISSED KEYS:");
        }
        for (String missedKey : missedKeys) {
            //out.println("\t" + missedKey);
            Integer index = keysIndex.get(missedKey);
            HSSFRow row = sheet.getRow(index);
            HSSFCell cell = row.createCell((int)row.getLastCellNum());
            cell.setCellStyle(missedStyle);
            boolean untranslatable = this.untranslatableMap.get(missedKey) != null && this.untranslatableMap.get(missedKey);
            if (!untranslatable) {
                out.println("\t" + missedKey);
            }
            cell.setCellStyle(untranslatable? untranstableStyle: missedStyle);
        }

        FileOutputStream outStream = new FileOutputStream(f);
        wb.write(outStream);
        outStream.close();

        if (missedKeys.isEmpty()) {
            out.println(String.format("'%s' was processed", lang));
        } else {
            out.println(String.format("'%s' was processed with MISSED KEYS - %d", lang, missedKeys.size()));
        }
    }
}
