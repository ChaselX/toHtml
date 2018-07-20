package com.example.office.utils;

import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PaneInformation;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;

import java.io.*;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class POIReadExcel {
//	private static Logger logger = Logger.getLogger(POIReadExcel.class);

    /**
     * 0保存合并单元格的对应起始和截止单元格、1保存被合并的那些单元格、2记录被隐藏的单元格个数、3记录合并了单元格，但是合并的首行被隐藏的情况
     */
    private static Map<String, Object> map[];

    /**
     * /home/local/nfs/RCSOA_UPLOAD_TEST/attachment/internalinformation_file/excelPic/
     */
    private static String UPLOAD_FILE;

    /**
     * http://117.136.240.85/RCSOA_UPLOAD_TEST/attachment/internalinformation_file/excelPic/
     */
    private static String READ_FILE;

    /**
     * excel转html输出最新（excel2Html.html）
     */
    private static String EXCEL_TO_HTML = "excel2Html.html";

    /**
     * 程序入口方法（读取指定位置的excel，将其转换成html形式的字符串，并保存成同名的html文件在相同的目录下，默认带样式）
     *
     * @return <table>...</table> 字符串
     */
    public static String excelWriteToHtml(String sourcePath) {
        File sourceFile = new File(sourcePath);
        try {
            InputStream fis = new FileInputStream(sourceFile);
            String excelHtml = POIReadExcel.readExcelToHtml(fis, true);
            return excelHtml;
        } catch (FileNotFoundException e) {
//        	logger.error("excelWriteToHtml error: "+e.getMessage(), e);
            return null;
        }
    }

    /**
     * 程序入口方法（将指定路径的excel文件读取成字符串）
     *
     * @param filePath    文件的路径
     * @param isWithStyle 是否需要表格样式 包含 字体 颜色 边框 对齐方式
     * @return <table>...</table> 字符串
     */
    public static String readExcelToHtml(String filePath, boolean isWithStyle) {
        InputStream is = null;
        String htmlExcel = null;
        try {
            File sourcefile = new File(filePath);
            is = new FileInputStream(sourcefile);
            Workbook wb = WorkbookFactory.create(is);
            htmlExcel = readWorkbook(wb, isWithStyle);
        } catch (Exception e) {
//        	logger.error("readExcelToHtml error: "+e.getMessage(), e);
        } finally {
            try {
                is.close();
            } catch (IOException e) {
//            	logger.error("readExcelToHtml error: "+e.getMessage(), e);
            }
        }
        return htmlExcel;
    }

    /**
     * 程序入口方法（将指定路径的excel文件读取成字符串）
     *
     * @param is          excel转换成的输入流
     * @param isWithStyle 是否需要表格样式 包含 字体 颜色 边框 对齐方式
     * @return <table>...</table> 字符串
     */
    public static String readExcelToHtml(InputStream is, boolean isWithStyle) {
        String htmlExcel = null;
        try {
            Workbook wb = WorkbookFactory.create(is);
            htmlExcel = readWorkbook(wb, isWithStyle);
        } catch (Exception e) {
//        	logger.error("readExcelToHtml error: "+e.getMessage(), e);
        } finally {
            try {
                is.close();
            } catch (IOException e) {
//            	logger.error("readExcelToHtml error: "+e.getMessage(), e);
            }
        }
        return htmlExcel;
    }

    /**
     * excel转换成带样式的html
     * <p>add by CJ 2018年5月19日</p>
     *
     * @param is      excel文件输入流
     * @param infoMap (上传路径/home，读取路径IP)
     * @return
     */
    public static String readExcelToHtml(InputStream is, Map<String, String> infoMap) {
        String htmlExcel = null;
        try {
            UPLOAD_FILE = infoMap.get("uploadFile");
            READ_FILE = infoMap.get("readfile");
//        	logger.info(String.format("1、readExcelToHtml uploadFile: %s%s", UPLOAD_FILE, DateUtils.getUIIDByCurrentTime()));
            Workbook wb = WorkbookFactory.create(is);
            htmlExcel = readWorkbook(wb, true);
            printExcel2Html(htmlExcel);
//            logger.info(String.format("2、readExcelToHtml uploadFile: %s%s", READ_FILE, DateUtils.getUIIDByCurrentTime()));
        } catch (Exception e) {
//        	logger.error("readExcelToHtml error: "+e.getMessage(), e);
        } finally {
            try {
                is.close();
            } catch (IOException e) {
//            	logger.error("readExcelToHtml error: "+e.getMessage(), e);
            }
        }
        return htmlExcel;
    }

    /**
     * 根据excel的版本分配不同的读取方法进行处理
     *
     * @param wb
     * @param isWithStyle
     * @return
     */
    private static String readWorkbook(Workbook wb, boolean isWithStyle) {
        String htmlExcel = "";
        if (wb instanceof XSSFWorkbook) {
            XSSFWorkbook xWb = (XSSFWorkbook) wb;
            htmlExcel = getExcelInfo(xWb, isWithStyle);
        } else if (wb instanceof HSSFWorkbook) {
            HSSFWorkbook hWb = (HSSFWorkbook) wb;
            htmlExcel = getExcelInfo(hWb, isWithStyle);
        }
        return htmlExcel;
    }

    /**
     * 读取excel成string
     *
     * @param wb
     * @param isWithStyle
     * @return
     */
    public static String getExcelInfo(Workbook wb, boolean isWithStyle) {
        StringBuffer sb = new StringBuffer();
        int sheetsNum = wb.getNumberOfSheets();


        String sheetNames = "";
        sb.append("<head>");
        sb.append("<style>body{\n" +
                "    margin:0;\n" +
                "    padding:0;\n" +
                "\n" +
                "}\n" +
                "#box{\n" +
                "    position: fixed;\n" +
                "    width:100%;\n" +
                "    height: 100%;\n" +
                "    z-index:99999;\n" +
                "    display: flex;\n" +
                "    flex-flow: column;\n" +
                "}\n" +
                "#sheet{\n" +
                "    width: 100%;\n" +
                "    display: flex;\n" +
                "    justify-content: space-around;\n" +
                "    height: 1.07rem;\n" +
                "}\n" +
                "#sheet span{\n" +
                "    flex: 1;\n" +
                "    display: block;\n" +
                "    text-align: center;\n" +
                "    cursor: pointer;\n" +
                "    line-height: 1.07rem;\n" +
                "    \n" +
                "}\n" +
                "#content{\n" +
                "    flex:1;\n" +
                "    width: 100%;\n" +
                "    height:100%;\n" +
                "    overflow:scroll;\n" +
                "    -webkit-overflow-scrolling: touch;\n" +
                "}\n" +
                "#content div{\n" +
                "    width: 100%;\n" +
                "    height: 100%;\n" +
                "    display: none;\n" +
                "}\n" +
                "\n" +
                "table thead tr td{\n" +
                "    border:1px solid black !important;\n" +
                "}\n" +
                "\n" +
                "table thead{\n" +
                "    background:#fff;\n" +
                "}\n" +
                "\n" +
                "td{\n" +
                "    border:1px solid black !important;\n" +
                "}\n" +
                "\n" +
                "#content div:first-child{\n" +
                "    display: block;\n" +
                "}\n" +
                "#sheet .active{\n" +
                "    background: #ddd;\n" +
                "    color: #ffffff;\n" +
                "}\n</style>");
        sb.append("<meta charset='UTF-8'></head><body>");
        for (int i = 0; i < sheetsNum; i++) {
            Sheet s = wb.getSheetAt(i);// 获取第一个Sheet的内容
            if (i == 0) {
                sheetNames += "<span class='active'>" + s.getSheetName() + "</span>";
            } else {
                sheetNames += "<span>" + s.getSheetName() + "</span>";
            }

        }
        sb.append("<div id='box'> <div id='sheet'>");
        sb.append(sheetNames);
        sb.append("</div><div id='content'>");
        for (int sheetNum = 0; sheetNum < sheetsNum; sheetNum++) {
            Sheet sheet = wb.getSheetAt(sheetNum);


            // map等待存储excel图片
            Map<String, PictureData> sheetIndexPicMap = getSheetPictrues(sheetNum, sheet, wb);
            Map<String, String> imgMap = new HashMap<String, String>();
            if (sheetIndexPicMap != null) {
                imgMap = printImg(sheetIndexPicMap);
                printImpToWb(imgMap, wb);
            }
            map = getRowSpanColSpanMap(sheet);

            PaneInformation information = sheet.getPaneInformation();
            short num = 0;
            sheetNames += sheet.getSheetName();
            if (information != null) {
                if (information.isFreezePane()) {
                    num = information.getHorizontalSplitPosition();
                    Row r = sheet.getRow(num - 1);
                    Cell c = r.getCell(num - 1);
                }
            }
            sb.append("<div><table style='margin-bottom:90%;' border='1' cellpadding='0' cellspacing=0>");
            //读取excel拼装html
            int lastRowNum = sheet.getLastRowNum();
            Row row = null;      //兼容
            Cell cell = null;    //兼容
            boolean flag = true;
            boolean headFlag = true;
            for (int rowNum = sheet.getFirstRowNum(); rowNum <= lastRowNum; rowNum++) {
                if (rowNum > 1000) break;
                row = sheet.getRow(rowNum);

                int lastColNum = POIReadExcel.getColsOfTable(sheet)[0];
                int rowHeight = POIReadExcel.getColsOfTable(sheet)[1];

                if (null != row) {
                    lastColNum = row.getLastCellNum();
                    rowHeight = row.getHeight();
                }
                if (null == row) {
                    sb.append("<tr><td >  </td></tr>");
                    continue;
                } else if (row.getZeroHeight()) {
                    continue;
                } else if (0 == rowHeight) {
                    continue;     //针对jxl的隐藏行（此类隐藏行只是把高度设置为0，单getZeroHeight无法识别）
                }

                if (num != 0 && rowNum <= num - 1) {
                    if (headFlag) {
                        sb.append("<thead><tr>");
                        headFlag = false;
                    } else {
                        sb.append("<tr>");
                    }

                } else {
                    if (flag) {
                        sb.append("<tr>");
                    } else {
                        sb.append("<tbody><tr>");
                        flag = false;
                    }

                }

                for (int colNum = 0; colNum < lastColNum; colNum++) {
                    if (sheet.isColumnHidden(colNum)) continue;
                    /* 各sheet.. */
                    String imageRowNum = sheetNum + "_" + rowNum + "_" + colNum;
                    String imageHtml = "";
                    cell = row.getCell(colNum);
                    //特殊情况 空白的单元格会返回null+//判断该单元格是否包含图片，为空时也可能包含图片
                    if ((sheetIndexPicMap != null && !sheetIndexPicMap.containsKey(imageRowNum) || sheetIndexPicMap == null) && cell == null) {
                        sb.append("<td>  </td>");
                        continue;
                    }
                    if (sheetIndexPicMap != null && sheetIndexPicMap.containsKey(imageRowNum)) {
                        String imagePath = imgMap.get(imageRowNum);
                        imageHtml = "<img src='" + imagePath + "' style='height:auto;'>";
                    }
                    String
                            stringValue = getCellValue(cell);
                    if (map[0].containsKey(rowNum + "," + colNum)) {
                        String pointString = (String) map[0].get(rowNum + "," + colNum);
                        int bottomeRow = Integer.valueOf(pointString.split(",")[0]);
                        int bottomeCol = Integer.valueOf(pointString.split(",")[1]);
                        int rowSpan = bottomeRow - rowNum + 1;
                        int colSpan = bottomeCol - colNum + 1;
                        if (map[2].containsKey(rowNum + "," + colNum)) {
                            rowSpan = rowSpan - (Integer) map[2].get(rowNum + "," + colNum);
                        }
                        sb.append("<td rowspan= '" + rowSpan + "' colspan= '" + colSpan + "' ");
                        if (map.length > 3 && map[3].containsKey(rowNum + "," + colNum)) {
                            //此类数据首行被隐藏，value为空，需使用其他方式获取值
                            stringValue = getMergedRegionValue(sheet, rowNum, colNum);
                        }
                    } else if (map[1].containsKey(rowNum + "," + colNum)) {
                        map[1].remove(rowNum + "," + colNum);
                        continue;
                    } else {
                        sb.append("<td ");
                    }

                    //判断是否需要样式
                    if (isWithStyle) {
                        if (sheetIndexPicMap != null && sheetIndexPicMap.containsKey(imageRowNum)) {
                            dealExcelStyle(wb, sheet, cell, sb, 1);
                        } else {
                            dealExcelStyle(wb, sheet, cell, sb, 0);//处理单元格样式
                        }
                    }
                    sb.append(">");

                    if (sheetIndexPicMap != null && sheetIndexPicMap.containsKey(imageRowNum)) {
                        sb.append(imageHtml);
                    }
                    if (stringValue == null || "".equals(stringValue.trim())) {
                        sb.append("   ");
                    } else {
                        // 将ascii码为160的空格转换为html下的空格（ ）
                        sb.append(stringValue.replace(String.valueOf((char) 160), " "));
                    }
                    sb.append("</td>");
                }
                if (num != 0 && rowNum <= num - 1) {
                    if (rowNum == num - 1) {
                        sb.append("</thead></tr>");
                    } else {
                        sb.append("</tr>");
                    }
                } else {
                    if (flag) {
                        sb.append("</tr>");
                    } else {
                        sb.append("</tr></tbody>");
                    }
                }
            }
            sb.append("</table></div>");

        }
        sb.append("</div>");
        return sb.toString();
    }

    /**
     * 分析excel表格，记录合并单元格相关的参数，用于之后html页面元素的合并操作
     *
     * @param sheet
     * @return
     */
    private static Map<String, Object>[] getRowSpanColSpanMap(Sheet sheet) {
        Map<String, String> map0 = new HashMap<String, String>();   //保存合并单元格的对应起始和截止单元格
        Map<String, String> map1 = new HashMap<String, String>();   //保存被合并的那些单元格
        Map<String, Integer> map2 = new HashMap<String, Integer>(); //记录被隐藏的单元格个数
        Map<String, String> map3 = new HashMap<String, String>();   //记录合并了单元格，但是合并的首行被隐藏的情况
        int mergedNum = sheet.getNumMergedRegions();
        CellRangeAddress range = null;
        Row row = null;
        for (int i = 0; i < mergedNum; i++) {
            range = sheet.getMergedRegion(i);
            int topRow = range.getFirstRow();
            int topCol = range.getFirstColumn();
            int bottomRow = range.getLastRow();
            int bottomCol = range.getLastColumn();
            /**
             * 此类数据为合并了单元格的数据
             * 1.处理隐藏（只处理行隐藏，列隐藏poi已经处理）
             */
            if (topRow != bottomRow) {
                int zeroRoleNum = 0;
                int tempRow = topRow;
                for (int j = topRow; j <= bottomRow; j++) {
                    row = sheet.getRow(j);
                    if (row.getZeroHeight() || row.getHeight() == 0) {
                        if (j == tempRow) {
                            //首行就进行隐藏，将rowTop向后移
                            tempRow++;
                            continue;//由于top下移，后面计算rowSpan时会扣除移走的列，所以不必增加zeroRoleNum;
                        }
                        zeroRoleNum++;
                    }
                }
                if (tempRow != topRow) {
                    map3.put(tempRow + "," + topCol, topRow + "," + topCol);
                    topRow = tempRow;
                }
                if (zeroRoleNum != 0) map2.put(topRow + "," + topCol, zeroRoleNum);
            }
            map0.put(topRow + "," + topCol, bottomRow + "," + bottomCol);
            int tempRow = topRow;
            while (tempRow <= bottomRow) {
                int tempCol = topCol;
                while (tempCol <= bottomCol) {
                    map1.put(tempRow + "," + tempCol, topRow + "," + topCol);
                    tempCol++;
                }
                tempRow++;
            }
            map1.remove(topRow + "," + topCol);
        }
        Map[] map = {map0, map1, map2, map3};
//        logger.info("getRowSpanColSpanMap :"+map0.toString());
        return map;
    }


    /**
     * 获取合并单元格的值
     *
     * @param sheet
     * @param row
     * @param column
     * @return
     */
    public static String getMergedRegionValue(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress ca = sheet.getMergedRegion(i);
            int firstColumn = ca.getFirstColumn();
            int lastColumn = ca.getLastColumn();
            int firstRow = ca.getFirstRow();
            int lastRow = ca.getLastRow();

            if (row >= firstRow && row <= lastRow) {

                if (column >= firstColumn && column <= lastColumn) {
                    Row fRow = sheet.getRow(firstRow);
                    Cell fCell = fRow.getCell(firstColumn);

                    return getCellValue(fCell);
                }
            }
        }
        return null;
    }

    /**
     * 获取表格单元格Cell内容
     *
     * @param cell
     * @return
     */
    private static String getCellValue(Cell cell) {
        String result = new String();
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_NUMERIC:// 数字类型
                if (HSSFDateUtil.isCellDateFormatted(cell)) {// 处理日期格式、时间格式
                    SimpleDateFormat sdf = null;
                    if (cell.getCellStyle().getDataFormat() == HSSFDataFormat.getBuiltinFormat("h:mm")) {
                        sdf = new SimpleDateFormat("HH:mm");
                    } else {// 日期
                        sdf = new SimpleDateFormat("yyyy-MM-dd");
                    }
                    Date date = cell.getDateCellValue();
                    result = sdf.format(date);
                } else if (cell.getCellStyle().getDataFormat() == 58) {
                    // 处理自定义日期格式：m月d日(通过判断单元格的格式id解决，id的值是58)
                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                    double value = cell.getNumericCellValue();
                    Date date = org.apache.poi.ss.usermodel.DateUtil
                            .getJavaDate(value);
                    result = sdf.format(date);
                } else {
                    double value = cell.getNumericCellValue();
                    CellStyle style = cell.getCellStyle();
                    DecimalFormat format = new DecimalFormat();
                    String formatStr = style.getDataFormatString();
                    // 单元格设置成常规
                    if (formatStr.equals("General")) {
                        format.applyPattern("#");
                    } else if (formatStr.contains("%")) {
                        format.applyPattern(formatStr);
                    }
                    result = format.format(value);
                }
                break;
            case Cell.CELL_TYPE_STRING:// String类型
                result = cell.getRichStringCellValue().toString();
                break;
            case Cell.CELL_TYPE_BLANK:
                result = "";
                break;
            case Cell.CELL_TYPE_FORMULA:
                switch (cell.getCachedFormulaResultType()) {
                    case Cell.CELL_TYPE_BOOLEAN:
                        result = String.valueOf(cell.getBooleanCellValue());
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        result = String.valueOf(cell.getNumericCellValue());
                        break;
                    case Cell.CELL_TYPE_STRING:
                        result = String.valueOf(cell.getStringCellValue());
                        break;
                    case Cell.CELL_TYPE_FORMULA:
                        break;
                    default:
                        result = "";
                        break;
                }
                break;
            default:
                result = "";
                break;
        }
        return result;
    }

    /**
     * 处理表格样式
     *
     * @param wb
     * @param sheet
     * @param cell
     * @param sb
     * @param flag  0文字；1图片
     */
    private static void dealExcelStyle(Workbook wb, Sheet sheet, Cell cell, StringBuffer sb, int flag) {
        CellStyle cellStyle = cell.getCellStyle();
        if (cellStyle != null) {
            HorizontalAlignment alignment = cellStyle.getAlignmentEnum();
            sb.append("align='" + convertAlignToHtml(alignment) + "' ");//单元格内容的水平对齐方式
            VerticalAlignment verticalAlignment = cellStyle.getVerticalAlignmentEnum();
            sb.append("valign='" + convertVerticalAlignToHtml(verticalAlignment) + "' ");//单元格中内容的垂直排列方式

            if (wb instanceof XSSFWorkbook) {

                XSSFFont xf = ((XSSFCellStyle) cellStyle).getFont();
                boolean boldWeight = xf.getBold();
                sb.append("style='");
                sb.append("font-weight:" + (boldWeight ? "bold" : "normal") + ";"); // 字体加粗
                sb.append("font-size: " + xf.getFontHeight() / 2 + "%;"); // 字体大小

                int topRow = cell.getRowIndex(), topColumn = cell.getColumnIndex();
                if (map[0].containsKey(topRow + "," + topColumn)) {//该单元格为合并单元格，宽度需要获取所有单元格宽度后合并
                    String value = (String) map[0].get(topRow + "," + topColumn);
                    String[] ary = value.split(",");
                    int bottomColumn = Integer.parseInt(ary[1]);
                    if (topColumn != bottomColumn) {//合并列，需要计算相应宽度
                        int columnWidth = 0;
                        for (int i = topColumn; i <= bottomColumn; i++) {
                            columnWidth += sheet.getColumnWidth(i);
                        }
                        sb.append("width:" + columnWidth / 256 * xf.getFontHeight() / 20 + "pt;");
                    } else {
                        int columnWidth = sheet.getColumnWidth(cell.getColumnIndex());
                        sb.append("width:" + columnWidth / 256 * xf.getFontHeight() / 20 + "pt;");
                    }
                } else {
                    int columnWidth = sheet.getColumnWidth(cell.getColumnIndex());
                    sb.append("width:" + columnWidth / 256 * xf.getFontHeight() / 20 + "pt;");
                }

                XSSFColor xc = xf.getXSSFColor();
                if (xc != null && !"".equals(xc.toString())) {
                    sb.append("color:#" + xc.getARGBHex().substring(2) + ";"); // 字体颜色
                }

                XSSFColor fgColor = (XSSFColor) cellStyle.getFillForegroundColorColor();
                if (fgColor != null && !"".equals(fgColor.toString())) {
                    sb.append("background-color:#" + fgColor.getARGBHex().substring(2) + ";"); // 背景颜色
                }
                /** 图片设置浮动 */
                if (flag == 1) {
                    sb.append("float:left;");
                } else {
                    /** 非图片设置边框样式、边框颜色、边框厚度 */
                    sb.append("border:solid #B8B8B8 1px;");
                }
            } else if (wb instanceof HSSFWorkbook) {
                HSSFFont hf = ((HSSFCellStyle) cellStyle).getFont(wb);
                boolean boldWeight = hf.getBold();
                short fontColor = hf.getColor();
                sb.append("style='");

                HSSFPalette palette = ((HSSFWorkbook) wb).getCustomPalette(); // 类HSSFPalette用于求的颜色的国际标准形式
                HSSFColor hc = palette.getColor(fontColor);
                sb.append("font-weight:" + (boldWeight ? "bold" : "normal") + ";"); // 字体加粗
                sb.append("font-size: " + hf.getFontHeight() / 2 + "%;"); // 字体大小
                String fontColorStr = convertToStardColor(hc);
                if (fontColorStr != null && !"".equals(fontColorStr.trim())) {
                    sb.append("color:" + fontColorStr + ";"); // 字体颜色
                }

                int topRow = cell.getRowIndex(), topColumn = cell.getColumnIndex();
                if (map[0].containsKey(topRow + "," + topColumn)) {//该单元格为合并单元格，宽度需要获取所有单元格宽度后合并
                    String value = (String) map[0].get(topRow + "," + topColumn);
                    String[] ary = value.split(",");
                    int bottomColumn = Integer.parseInt(ary[1]);
                    if (topColumn != bottomColumn) {//合并列，需要计算相应宽度
                        int columnWidth = 0;
                        for (int i = topColumn; i <= bottomColumn; i++) {
                            columnWidth += sheet.getColumnWidth(i);
                        }
                        sb.append("width:" + columnWidth / 256 * hf.getFontHeight() / 20 + "pt;");
                    } else {
                        int columnWidth = sheet.getColumnWidth(cell.getColumnIndex());
                        sb.append("width:" + columnWidth / 256 * hf.getFontHeight() / 20 + "pt;");
                    }
                } else {
                    int columnWidth = sheet.getColumnWidth(cell.getColumnIndex());
                    sb.append("width:" + columnWidth / 256 * hf.getFontHeight() / 20 + "pt;");
                }

                short bgColor = cellStyle.getFillForegroundColor();
                hc = palette.getColor(bgColor);
                String bgColorStr = convertToStardColor(hc);
                if (bgColorStr != null && !"".equals(bgColorStr.trim())) {
                    sb.append("background-color:" + bgColorStr + ";");      // 背景颜色
                }
                /** 图片设置浮动 */
                if (flag == 1) {
                    sb.append("float:left;");
                    /** 非图片设置边框样式、边框颜色、边框厚度 */
                } else {
                    sb.append("border:solid #B8B8B8 1px;");
                }
            }
            sb.append("' ");
        }
    }

    /**
     * 单元格内容的水平对齐方式
     *
     * @param alignment
     * @return
     */
    private static String convertAlignToHtml(HorizontalAlignment alignment) {
        String align = "left";
        switch (alignment) {
            case LEFT:
                align = "left";
                break;
            case CENTER:
                align = "center";
                break;
            case RIGHT:
                align = "right";
                break;
            default:
                break;
        }
        return align;
    }

    /**
     * 单元格中内容的垂直排列方式
     *
     * @param verticalAlignment
     * @return
     */
    private static String convertVerticalAlignToHtml(VerticalAlignment verticalAlignment) {
        String valign = "middle";
        switch (verticalAlignment) {
            case BOTTOM:
                valign = "bottom";
                break;
            case CENTER:
                valign = "center";
                break;
            case TOP:
                valign = "top";
                break;
            default:
                break;
        }
        return valign;
    }

    private static String convertToStardColor(HSSFColor hc) {
        StringBuffer sb = new StringBuffer("");
        if (hc != null) {
            if (HSSFColor.AUTOMATIC.index == hc.getIndex()) {
                return null;
            }
            sb.append("#");
            for (int i = 0; i < hc.getTriplet().length; i++) {
                sb.append(fillWithZero(Integer.toHexString(hc.getTriplet()[i])));
            }
        }
        return sb.toString();
    }

    private static String fillWithZero(String str) {
        if (str != null && str.length() < 2) {
            return "0" + str;
        }
        return str;
    }

    static String[] bordesr = {"border-top:", "border-right:", "border-bottom:", "border-left:"};
    static String[] borderStyles = {"solid ", "solid ", "solid ", "solid ", "solid ", "solid ", "solid ", "solid ", "solid ", "solid", "solid", "solid", "solid", "solid"};

    @SuppressWarnings("unused")
    private static String getBorderStyle(HSSFPalette palette, int b, short s, short t) {
        if (s == 0) return bordesr[b] + borderStyles[s] + "#d0d7e5 1px;";
        String borderColorStr = convertToStardColor(palette.getColor(t));
        borderColorStr = borderColorStr == null || borderColorStr.length() < 1 ? "#000000" : borderColorStr;
        return bordesr[b] + borderStyles[s] + borderColorStr + " 1px;";
    }

    @SuppressWarnings("unused")
    private static String getBorderStyle(int b, short s, XSSFColor xc) {
        if (s == 0) return bordesr[b] + borderStyles[s] + "#d0d7e5 1px;";
        if (xc != null && !"".equals(xc)) {
            String borderColorStr = xc.getARGBHex();//t.getARGBHex();
            borderColorStr = borderColorStr == null || borderColorStr.length() < 1 ? "#000000" : borderColorStr.substring(2);
            return bordesr[b] + borderStyles[s] + borderColorStr + " 1px;";
        }
        return "";
    }

    @SuppressWarnings("unused")
    private static void writeFile(String content, String path) {
        OutputStream os = null;
        BufferedWriter bw = null;
        try {
            File file = new File(path);
            os = new FileOutputStream(file);
            bw = new BufferedWriter(new OutputStreamWriter(os, "UTF-8"));
            bw.write(content);
        } catch (FileNotFoundException e) {
//        	logger.error("writeFile error: "+e.getMessage(), e);
        } catch (IOException e) {
//        	logger.error("writeFile error: "+e.getMessage(), e);
        } finally {
            try {
                if (null != bw)
                    bw.close();
                if (null != os)
                    os.close();
            } catch (IOException e) {
//            	logger.error("writeFile error: "+e.getMessage(), e);
            }
        }
    }

    /**
     * 获取Excel图片公共方法
     *
     * @param sheetNum 当前sheet编号
     * @param sheet    当前sheet对象
     * @param workbook 工作簿对象
     * @return Map key:图片单元格索引（0_1_1）String，value:图片流PictureData
     */
    public static Map<String, PictureData> getSheetPictrues(int sheetNum, Sheet sheet, Workbook workbook) {
        if (workbook instanceof HSSFWorkbook) {
            return getSheetPictrues03(sheetNum, (HSSFSheet) sheet, (HSSFWorkbook) workbook);
        } else if (workbook instanceof XSSFWorkbook) {
            return getSheetPictrues07(sheetNum, (XSSFSheet) sheet, (XSSFWorkbook) workbook);
        } else {
            return null;
        }
    }

    /**
     * 获取Excel2003图片
     *
     * @param sheetNum 当前sheet编号
     * @param sheet    当前sheet对象
     * @param workbook 工作簿对象
     * @return Map key:图片单元格索引（0_1_1）String，value:图片流PictureData
     * @throws IOException
     */
    private static Map<String, PictureData> getSheetPictrues03(int sheetNum,
                                                               HSSFSheet sheet, HSSFWorkbook workbook) {
        Map<String, PictureData> sheetIndexPicMap = new HashMap<String, PictureData>();
        List<HSSFPictureData> pictures = workbook.getAllPictures();
        if (pictures.size() != 0) {
            for (HSSFShape shape : sheet.getDrawingPatriarch().getChildren()) {
                HSSFClientAnchor anchor = (HSSFClientAnchor) shape.getAnchor();
                shape.getLineWidth();
                if (shape instanceof HSSFPicture) {
                    HSSFPicture pic = (HSSFPicture) shape;
                    int pictureIndex = pic.getPictureIndex() - 1;
                    HSSFPictureData picData = pictures.get(pictureIndex);
                    String picIndex = String.valueOf(sheetNum) + "_"
                            + String.valueOf(anchor.getRow1()) + "_"
                            + String.valueOf(anchor.getCol1());
                    sheetIndexPicMap.put(picIndex, picData);
                }
            }
            return sheetIndexPicMap;
        } else {
            return null;
        }
    }

    /**
     * 获取Excel2007图片
     *
     * @param sheetNum 当前sheet编号
     * @param sheet    当前sheet对象
     * @param workbook 工作簿对象
     * @return Map key:图片单元格索引（0_1_1）String，value:图片流PictureData
     */
    private static Map<String, PictureData> getSheetPictrues07(int sheetNum,
                                                               XSSFSheet sheet, XSSFWorkbook workbook) {
        Map<String, PictureData> sheetIndexPicMap = new HashMap<String, PictureData>();
        for (POIXMLDocumentPart dr : sheet.getRelations()) {
            if (dr instanceof XSSFDrawing) {
                XSSFDrawing drawing = (XSSFDrawing) dr;
                List<XSSFShape> shapes = drawing.getShapes();
                for (XSSFShape shape : shapes) {
                    XSSFPicture pic = (XSSFPicture) shape;
                    XSSFClientAnchor anchor = pic.getPreferredSize();
                    CTMarker ctMarker = anchor.getFrom();
                    String picIndex = String.valueOf(sheetNum) + "_"
                            + ctMarker.getRow() + "_"
                            + ctMarker.getCol();
                    sheetIndexPicMap.put(picIndex, pic.getPictureData());
                }
            }
        }
        return sheetIndexPicMap;
    }

    public static void printImg(List<Map<String, PictureData>> sheetList) throws IOException {
        for (Map<String, PictureData> map : sheetList) {
            printImg(map);
        }
    }

    /**
     * 写入图片到FastDFS
     * <p>add by CJ 2018年5月19日</p>
     *
     * @param map
     * @return
     */
    public static Map<String, String> printImg(Map<String, PictureData> map) {
        Map<String, String> imgMap = new HashMap<String, String>();
        String imgName = null;
        try {
            Object key[] = map.keySet().toArray();
            for (int i = 0; i < map.size(); i++) {
                // 获取图片流
                PictureData pic = map.get(key[i]);
                // 获取图片索引
                String picName = key[i].toString();
                // 获取图片格式
                String ext = pic.suggestFileExtension();
                byte[] data = pic.getData();
                File uploadFile = new File(UPLOAD_FILE);
                if (!uploadFile.exists()) {
                    uploadFile.mkdirs();
                }
//    			imgName = picName + "_" + DateUtils.getUIIDByCurrentTime() + "." + ext;
                imgName = picName + "_" + new Date().getTime() + "." + ext;
                FileOutputStream out = new FileOutputStream(UPLOAD_FILE + imgName);
                imgMap.put(picName, READ_FILE + imgName);
                out.write(data);
                out.flush();
                out.close();
            }
        } catch (Exception e) {
//			logger.error("printImg error: "+e.getMessage(), e);
        }
        return imgMap;
    }

    /**
     * 对图片单元格赋值使其可读取到
     * <p>add by CJ 2018年5月21日</p>
     *
     * @param imgMap
     * @param wb
     */
    @SuppressWarnings("unused")
    private static void printImpToWb(Map<String, String> imgMap, Workbook wb) {
        Sheet sheet = null;
        Row row = null;
        String[] sheetRowCol = new String[3];
        for (String key : imgMap.keySet()) {
            sheetRowCol = key.split("_");
            sheet = wb.getSheetAt(Integer.parseInt(sheetRowCol[0]));
            row = sheet.getRow(Integer.parseInt(sheetRowCol[1])) == null ? sheet.createRow(Integer.parseInt(sheetRowCol[1])) :
                    sheet.getRow(Integer.parseInt(sheetRowCol[1]));
            Cell cell = row.getCell(Integer.parseInt(sheetRowCol[2])) == null ? row.createCell(Integer.parseInt(sheetRowCol[2])) :
                    row.getCell(Integer.parseInt(sheetRowCol[2]));
            /* 设置行高 row.height? */
        }
    }

    private static int[] getColsOfTable(Sheet sheet) {
        int[] data = {0, 0};
        for (int i = sheet.getFirstRowNum(); i < sheet.getLastRowNum(); i++) {
            if (null != sheet.getRow(i)) {
                data[0] = sheet.getRow(i).getLastCellNum();
                data[1] = sheet.getRow(i).getHeight();
            } else
                continue;
        }
        return data;
    }

    /**
     * 上传html到FastDFS
     * <p>add by CJ 2018年5月20日</p>
     *
     * @param html
     */
    public static void printExcel2Html(String html) {
        String htmlTempPath = null;
        try {
            htmlTempPath = UPLOAD_FILE + "HTML/";
            File uploadFile = new File(htmlTempPath);
            if (!uploadFile.exists()) {
                uploadFile.mkdirs();
            }
            FileOutputStream out = new FileOutputStream(htmlTempPath + EXCEL_TO_HTML);
            out.write(html.getBytes());
            out.flush();
            out.close();
        } catch (Exception e) {
//			logger.error("printImg error: "+e.getMessage(), e);
        }
    }
}
