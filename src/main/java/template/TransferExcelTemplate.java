package template;

import in.excel.po.TransferItem;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import po.Rmb;

import java.io.File;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * 业务回单模板
 */
public class TransferExcelTemplate implements ExcelTemplate{

    /**
     * 输出路径
     */
    private String output;

    private HSSFWorkbook wb;

    private HSSFSheet sheet;

    private HSSFPrintSetup hps;

    /**
     * 当前行的index
     */
    private int currentRowIndex;

    private Map<String, HSSFCellStyle> styleMap;

    private List<TransferItem> itemList;

    private String year;

    public TransferExcelTemplate(List<TransferItem> itemList, String output, String year, String month) {
        this.itemList = itemList;
        this.output = output;
        this.currentRowIndex = 0;
        //创建HSSFWorkbook对象
        this.wb = new HSSFWorkbook();
        this.year = year;

        // 初始化sheet
        initSheet(year, month);
        // 初始化样式
        initStyle();

        initPrintSetup();
    }
    @Override
    public void write() throws Exception {
        String dataFormat = new SimpleDateFormat("yyyyMMddHHmmss").format(new Date());
        File outFile = new File(output + "/业务回单_" + dataFormat + ".xls");
        FileOutputStream fos = new FileOutputStream(outFile);
        itemList.stream().forEach(item -> writeItem(item));

        wb.write(fos);
        System.out.println("生成[" + itemList.size() + "]条业务回单！路径：" + outFile.getAbsolutePath());
        wb.close();
        fos.close();
    }

    private void initSheet(String year, String month) {
        //创建HSSFSheet对象
        sheet = wb.createSheet(year+ "年" + month + "月");
        sheet.setColumnWidth(0, 1500);
        sheet.setColumnWidth(1, 2500);
        sheet.setColumnWidth(2, 8550);
        sheet.setColumnWidth(3, 4500);
        sheet.setColumnWidth(4, 1600);
        sheet.setColumnWidth(5, 3450);
        sheet.setColumnWidth(6, 3450);
    }

    private void initStyle() {
        styleMap = new HashMap<>();

        HSSFCellStyle title_style = wb.createCellStyle();
        title_style.setAlignment(HorizontalAlignment.RIGHT); // 垂直居中
        HSSFFont title_font = wb.createFont();
        title_font.setFontName("黑体"); //字体
        title_font.setFontHeightInPoints((short)20); //字体大小
        title_font.setBold(true); //字体加粗
        title_style.setVerticalAlignment(VerticalAlignment.CENTER); // 垂直居中
        title_style.setFont(title_font);
        styleMap.put("title_style", title_style);

        HSSFCellStyle small_title_style = wb.createCellStyle();
        HSSFFont small_title_font = wb.createFont();
        small_title_font.setFontName("宋体"); //字体
        small_title_font.setFontHeightInPoints((short)12); //字体大小
        small_title_style.setFont(small_title_font);
        small_title_style.setVerticalAlignment(VerticalAlignment.CENTER); // 垂直居中
        styleMap.put("small_title_style", small_title_style);

        HSSFCellStyle common_style = wb.createCellStyle();
        HSSFFont common_font = wb.createFont();
        common_font.setFontName("宋体"); //字体
        common_font.setFontHeightInPoints((short)11); //字体大小
        common_style.setFont(common_font);
        styleMap.put("common_style", common_style);

        HSSFCellStyle account_style = wb.createCellStyle();
        HSSFFont account_font = wb.createFont();
        account_font.setFontName("宋体"); //字体
        account_font.setFontHeightInPoints((short)11); //字体大小
        account_style.setFont(account_font);
        account_style.setWrapText(true);
        account_style.setVerticalAlignment(VerticalAlignment.TOP); // 垂直居中
        styleMap.put("account_style", account_style);

        HSSFCellStyle moneyStyle = wb.createCellStyle();
        moneyStyle.setAlignment(HorizontalAlignment.CENTER); // 水平居中
        HSSFFont font_SimSun_12 = wb.createFont();
        font_SimSun_12.setFontName("宋体"); //字体
        font_SimSun_12.setFontHeightInPoints((short)11); //字体大小
        small_title_style.setFont(small_title_font);
        moneyStyle.setFont(font_SimSun_12);
        moneyStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("#,##0.00"));
//        setBorder(moneyStyle);
        styleMap.put("moneyStyle", moneyStyle);
    }

    private void setBorder(HSSFCellStyle style) {
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
    }

    private void initPrintSetup() {
        this.hps = sheet.getPrintSetup();
        hps.setPaperSize(HSSFPrintSetup.A4_PAPERSIZE); // 纸张
        hps.setHeaderMargin(0.4);
        hps.setFooterMargin(0.4);
        sheet.setMargin(HSSFSheet.TopMargin, 0.4 ); // 上边距
        sheet.setMargin(HSSFSheet.BottomMargin, 0.3 ); // 下边距
        sheet.setMargin(HSSFSheet.LeftMargin, 0.9 ); // 左边距
        sheet.setMargin(HSSFSheet.RightMargin, 1 ); // 右边距
    }

    private void writeItem(TransferItem item) {
        float lineHeight = 32f;
        float commonLineHeight = 15f;
        HSSFRow row_1 = sheet.createRow(currentRowIndex);
        row_1.setHeightInPoints(lineHeight);
        HSSFCell cell_0_5 = row_1.createCell(0);
        cell_0_5.setCellValue("                                 业务回单");
        cell_0_5.setCellStyle(styleMap.get("title_style"));
        sheet.addMergedRegion(new CellRangeAddress(currentRowIndex, currentRowIndex,0,5));
        HSSFCell cell_0_6 = row_1.createCell(6);
        cell_0_6.setCellValue("（收款）");
        cell_0_6.setCellStyle(styleMap.get("small_title_style"));

        HSSFRow row_2 = sheet.createRow(currentRowIndex + 1);
        row_2.setHeightInPoints(commonLineHeight);
        HSSFCell cell_1_6 = row_2.createCell(0);
        cell_1_6.setCellValue("日期：" + year + "年" + item.getMonth() + "月" + item.getDay() + "日");
        cell_1_6.setCellStyle(styleMap.get("common_style"));
        sheet.addMergedRegion(new CellRangeAddress(currentRowIndex + 1, currentRowIndex + 1,0,6));

        HSSFRow row_3 = sheet.createRow(currentRowIndex + 2);
        row_3.setHeightInPoints(commonLineHeight);
        HSSFCell cell_2_6 = row_3.createCell(0);
        cell_2_6.setCellValue("回单编号：" + item.getTransferNum());
        cell_2_6.setCellStyle(styleMap.get("common_style"));
        sheet.addMergedRegion(new CellRangeAddress(currentRowIndex + 2, currentRowIndex + 2,0,2));

        HSSFRow row_4 = sheet.createRow(currentRowIndex + 3);
        row_4.setHeightInPoints(commonLineHeight);
        HSSFCell cell_3_6 = row_4.createCell(0);
        cell_3_6.setCellValue("付款人户名：" + item.getPayAccount());
        cell_3_6.setCellStyle(styleMap.get("common_style"));
        HSSFCell cell_3_4 = row_4.createCell(4);
        cell_3_4.setCellValue("付款人开户行：" + item.getPayBank());
        cell_3_4.setCellStyle(styleMap.get("account_style"));
        sheet.addMergedRegion(new CellRangeAddress(currentRowIndex + 3, currentRowIndex + 4,4,6));

        HSSFRow row_5 = sheet.createRow(currentRowIndex + 4);
        row_5.setHeightInPoints(commonLineHeight);
        HSSFCell cell_4_6 = row_5.createCell(0);
        cell_4_6.setCellValue("付款人账号(卡号)：" + item.getPayCard());
        cell_4_6.setCellStyle(styleMap.get("common_style"));

        HSSFRow row_6 = sheet.createRow(currentRowIndex + 5);
        row_6.setHeightInPoints(commonLineHeight);
        HSSFCell cell_5_6 = row_6.createCell(0);
        cell_5_6.setCellValue("收款人户名：四川东柳醪糟有限责任公司");
        cell_5_6.setCellStyle(styleMap.get("common_style"));
        HSSFCell cell_5_4 = row_6.createCell(4);
        cell_5_4.setCellValue("收款人开户行：达州大竹车坝支行");
        cell_5_4.setCellStyle(styleMap.get("account_style"));
        sheet.addMergedRegion(new CellRangeAddress(currentRowIndex + 5, currentRowIndex + 6,4,6));

        HSSFRow row_7 = sheet.createRow(currentRowIndex + 6);
        row_7.setHeightInPoints(commonLineHeight);
        HSSFCell cell_6_6 = row_7.createCell(0);
        cell_6_6.setCellStyle(styleMap.get("common_style"));
        cell_6_6.setCellValue("收款人账号（卡号）:2317008519020101048");

        double price = item.getPrice();
        HSSFRow row_8 = sheet.createRow(currentRowIndex + 7);
        row_8.setHeightInPoints(commonLineHeight);
        HSSFCell cell_7_6 = row_8.createCell(0);
        cell_7_6.setCellValue("金额：");
        cell_7_6.setCellStyle(styleMap.get("common_style"));
        HSSFCell cell_7_1 = row_8.createCell(1);
        cell_7_1.setCellValue(new Rmb(price).toHanStr());
        cell_7_1.setCellStyle(styleMap.get("common_style"));
        sheet.addMergedRegion(new CellRangeAddress(currentRowIndex + 7, currentRowIndex + 7,1,2));
        HSSFCell cell_7_4 = row_8.createCell(4);
        cell_7_4.setCellValue("小写：");
        cell_7_4.setCellStyle(styleMap.get("common_style"));
        HSSFCell cell_7_5 = row_8.createCell(5);
        cell_7_5.setCellValue(price);
        cell_7_5.setCellStyle(styleMap.get("moneyStyle"));
        HSSFCell cell_7_6_ = row_8.createCell(6);
        cell_7_6_.setCellValue("元");
        cell_7_6_.setCellStyle(styleMap.get("common_style"));

        HSSFRow row_9 = sheet.createRow(currentRowIndex + 8);
        row_9.setHeightInPoints(commonLineHeight);
        HSSFCell cell_8_6 = row_9.createCell(0);
        cell_8_6.setCellValue("业务 (产品) 种类：" + item.getBusinessType() + " 　　 凭证种类：000000000 ");
        cell_8_6.setCellStyle(styleMap.get("common_style"));
        HSSFCell cell_8_4 = row_9.createCell(4);
        cell_8_4.setCellValue("凭证号码：000000000000000000");
        cell_8_4.setCellStyle(styleMap.get("common_style"));

        HSSFRow row_10 = sheet.createRow(currentRowIndex + 9);
        row_10.setHeightInPoints(commonLineHeight);
        HSSFCell cell_9_2 = row_10.createCell(0);
        cell_9_2.setCellValue("摘要：" + item.getTips() + "    用途：               ");
        cell_9_2.setCellStyle(styleMap.get("common_style"));
        sheet.addMergedRegion(new CellRangeAddress(currentRowIndex + 9, currentRowIndex + 9,0,2));
        HSSFCell cell_9_4 = row_10.createCell(4);
        cell_9_4.setCellValue("币种：人民币");
        cell_9_4.setCellStyle(styleMap.get("common_style"));

        HSSFRow row_11 = sheet.createRow(currentRowIndex + 10);
        row_11.setHeightInPoints(commonLineHeight);
        int accountUserId = (int)(Math.random() * 900) + 100;
        HSSFCell cell_11_2 = row_11.createCell(0);
        cell_11_2.setCellValue("交易机构：0231700085        记账柜员：00" + accountUserId + "　 　交易代码：" + item.getTradeCode());
        cell_11_2.setCellStyle(styleMap.get("common_style"));
        HSSFCell cell_11_4 = row_11.createCell(4);
        cell_11_4.setCellValue("渠道：其他");
        cell_11_4.setCellStyle(styleMap.get("common_style"));

        HSSFRow row_12 = sheet.createRow(currentRowIndex + 11);
        row_12.setHeightInPoints(commonLineHeight);

        HSSFRow row_13 = sheet.createRow(currentRowIndex + 12);
        row_13.setHeightInPoints(commonLineHeight);
        HSSFCell cell_13_2 = row_13.createCell(0);
        cell_13_2.setCellValue("附言:货款");
        cell_13_2.setCellStyle(styleMap.get("common_style"));
        sheet.addMergedRegion(new CellRangeAddress(currentRowIndex + 12, currentRowIndex + 12,0,2));

        HSSFRow row_14 = sheet.createRow(currentRowIndex + 13);
        row_14.setHeightInPoints(commonLineHeight);
        HSSFCell cell_14_6 = row_14.createCell(0);
        cell_14_6.setCellValue("支付交易序号:" + item.getPayTradeCode() + " 报文种类:IBP101网银贷记业务报文  委托日期:" + year + "-" + item.getMonth() + "-" + item.getDay());
        cell_14_6.setCellStyle(styleMap.get("common_style"));
        sheet.addMergedRegion(new CellRangeAddress(currentRowIndex + 13, currentRowIndex + 13,0,6));

        HSSFRow row_15 = sheet.createRow(currentRowIndex + 14);
        row_15.setHeightInPoints(commonLineHeight);
        HSSFCell cell_15_6 = row_15.createCell(0);
        cell_15_6.setCellValue("业务种类:其他");
        cell_15_6.setCellStyle(styleMap.get("common_style"));

        HSSFRow row_16 = sheet.createRow(currentRowIndex + 15);
        row_16.setHeightInPoints(commonLineHeight);

        HSSFRow row_17 = sheet.createRow(currentRowIndex + 16);
        row_17.setHeightInPoints(commonLineHeight);
        HSSFCell cell_17_6 = row_17.createCell(0);
        Calendar calendar = Calendar.getInstance();
        calendar.set(Calendar.YEAR, Integer.valueOf(year));
        calendar.set(Calendar.MONTH, Integer.valueOf(item.getMonth()) - 1);
        calendar.set(Calendar.DAY_OF_MONTH, 2);
        calendar.add(Calendar.MONTH, 1);
        cell_17_6.setCellValue("本回单为第1次打印,注意重复      打印日期：" + calendar.get(Calendar.YEAR) + "年" + (calendar.get(Calendar.MONTH) + 1) + "月" + calendar.get(Calendar.DAY_OF_MONTH) + "日    打印柜员：9   ");
        cell_17_6.setCellStyle(styleMap.get("common_style"));

        HSSFRow row_18 = sheet.createRow(currentRowIndex + 17);
        row_18.setHeightInPoints(commonLineHeight);
        currentRowIndex += 18;
    }
}
