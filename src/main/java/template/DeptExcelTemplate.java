package template;

import cn.hutool.core.collection.CollectionUtil;
import in.excel.po.DebtItem;
import in.excel.po.Item;
import org.apache.commons.collections4.ListUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import po.Rmb;

import java.io.File;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class DeptExcelTemplate implements ExcelTemplate {

    private List<DebtItem> debts;

    private HSSFWorkbook wb;

    private HSSFSheet sheet;

    private String output;
    /**
     * 每张回拨单行数
     */
    private int deptRowCount;
    /**
     * 当前行的index
     */
    private int currentRowIndex;
    /**
     * 每张回拨单物展示品数量
     */
    private int itemCountOfDebt;

    private int totalRecords;

    private String year;

    private String month;

    private String day;

    private Map<String, HSSFCellStyle> styleMap;

    public DeptExcelTemplate(String output, List<DebtItem> debts, String year, String month, String day) {
        if (CollectionUtil.isEmpty(debts)) {
            throw new IllegalStateException("没有加载到数据！！！");
        }
        this.output = output;
        this.debts = debts;
        this.currentRowIndex = 0;
        this.deptRowCount = 15;
        this.itemCountOfDebt = 5;
        this.year = year;
        this.month = month;
        this.day = day;

        //创建HSSFWorkbook对象
        this.wb = new HSSFWorkbook();

        styleMap = new HashMap<>();
        // 初始化sheet
        this.initSheet();

        initStyle();
    }

    @Override
    public void write() throws Exception {
        for (DebtItem debt : debts) {
            writeDept(debt);
        }
        File file = new File(output);
        if (!file.exists()) {
            file.mkdir();
        }
        String dataFormat = new SimpleDateFormat("yyyyMMddhhmmss").format(new Date());
        File outFile = new File(output + "/回单记录_" + dataFormat + ".xls");
        FileOutputStream fos = new FileOutputStream(outFile);
        wb.write(fos);
        System.out.println("生成[" + totalRecords + "]条汇款单记录！路径：" + outFile.getAbsolutePath());
        fos.close();
    }

    private void writeDept(DebtItem debt) {
        List<Item> list = debt.getItems();
        List<List<Item>> lists = ListUtils.partition(list, itemCountOfDebt);
        if (lists.size() == 0) {
            return;
        }
        int count = lists.size();
        for (int i = 0;i < count;i++) {
            writeItem(debt, lists.get(i));
            totalRecords++;
        }
    }

    private void writeItem(DebtItem debt, List<Item> items) {
        int lineHeight = 20;


        HSSFRow row_1 = sheet.createRow(currentRowIndex);
        row_1.setHeightInPoints(lineHeight);
        HSSFRow row_1_2 = sheet.createRow(currentRowIndex + 1);
        row_1_2.setHeightInPoints(lineHeight);
        HSSFCell cell_0_16 = row_1.createCell(0);
        cell_0_16.setCellValue("四川东柳醪糟有限责任公司产 品 调 拨 单");
        HSSFCellStyle cell_0_16_style = wb.createCellStyle();
        cell_0_16_style.setAlignment(HorizontalAlignment.CENTER); // 水平居中
        cell_0_16_style.setVerticalAlignment(VerticalAlignment.CENTER); // 垂直居中
        HSSFFont cell_0_16_font = wb.createFont();
        cell_0_16_font.setFontName("黑体"); //字体
        cell_0_16_font.setFontHeightInPoints((short)14); //字体大小
        cell_0_16_style.setFont(cell_0_16_font);
        cell_0_16.setCellStyle(cell_0_16_style);
        sheet.createRow(currentRowIndex + 1);
        sheet.addMergedRegion(new CellRangeAddress(currentRowIndex, currentRowIndex + 1,0,15));

        HSSFRow row_2 = sheet.createRow(currentRowIndex + 2);
        row_2.setHeightInPoints(lineHeight);
        sheet.addMergedRegion(new CellRangeAddress(currentRowIndex + 2, currentRowIndex + 10,0,0));
        HSSFCell cell_10_0 = row_2.createCell(0);
        cell_10_0.setCellValue("︵东醪司︶财务部制");
        cell_10_0.setCellStyle(styleMap.get("style_SimSun_10"));
        HSSFCell cell_2_1 = row_2.createCell(1);
        cell_2_1.setCellValue("单位名称：");
        HSSFCellStyle cell_2_1_style = wb.createCellStyle();
        HSSFFont font_cell_2_1 = wb.createFont();
        font_cell_2_1.setFontName("宋体"); //字体
        font_cell_2_1.setFontHeightInPoints((short)12); //字体大小
        font_cell_2_1.setBold(true);
        cell_2_1_style.setFont(font_cell_2_1);
        cell_2_1_style.setVerticalAlignment(VerticalAlignment.CENTER); // 垂直居中
        cell_2_1.setCellStyle(cell_2_1_style);
        HSSFCell cell_2_2 = row_2.createCell(2);
        cell_2_2.setCellValue(debt.getCustom());
        sheet.addMergedRegion(new CellRangeAddress(currentRowIndex + 2, currentRowIndex + 2,2,7));
        cell_2_2.setCellStyle(styleMap.get("common_style"));
        HSSFCell cell_2_8 = row_2.createCell(8);
        cell_2_8.setCellValue(year + "年");
        cell_2_8.setCellStyle(styleMap.get("kaiti_style"));
        HSSFCell cell_2_9 = row_2.createCell(9);
        cell_2_9.setCellValue(month);
        cell_2_9.setCellStyle(styleMap.get("kaiti_center_12_style"));
        HSSFCell cell_2_10 = row_2.createCell(10);
        cell_2_10.setCellValue("月");
        cell_2_10.setCellStyle(styleMap.get("kaiti_style"));
        HSSFCell cell_2_11 = row_2.createCell(11);
        cell_2_11.setCellValue(day);
        cell_2_11.setCellStyle(styleMap.get("kaiti_center_12_style"));
        HSSFCell cell_2_12 = row_2.createCell(12);
        cell_2_12.setCellValue("日");
        cell_2_12.setCellStyle(styleMap.get("kaiti_style"));
        HSSFCell cell_2_13 = row_2.createCell(13);
        cell_2_13.setCellValue("编号：");
        cell_2_13.setCellStyle(styleMap.get("common_style"));
        HSSFCell cell_2_14 = row_2.createCell(14);
        cell_2_14.setCellValue("L2001001");
        cell_2_14.setCellStyle(styleMap.get("common_style"));
        HSSFCell cell_2_15 = row_2.createCell(15);
        sheet.addMergedRegion(new CellRangeAddress(currentRowIndex + 2, currentRowIndex + 10,15,15));
        cell_2_15.setCellValue("二 发货单位记账联");
        cell_2_15.setCellStyle(styleMap.get("style_SimSun_10"));

        HSSFRow row_3 = sheet.createRow(currentRowIndex + 3);
        row_3.setHeightInPoints(lineHeight);
        HSSFCell cell_3_1 = row_3.createCell(1);
        cell_3_1.setCellValue("产品名称");
        cell_3_1.setCellStyle(styleMap.get("kaiti_center_12_style"));
        sheet.addMergedRegion(new CellRangeAddress(currentRowIndex + 3, currentRowIndex + 3,1,2));
        HSSFCell cell_3_3 = row_3.createCell(3);
        cell_3_3.setCellValue("单位");
        cell_3_3.setCellStyle(styleMap.get("kaiti_center_12_style"));
        HSSFCell cell_3_4 = row_3.createCell(4);
        cell_3_4.setCellValue("总发数");
        cell_3_4.setCellStyle(styleMap.get("kaiti_center_12_style"));
        HSSFCell cell_3_5 = row_3.createCell(5);
        cell_3_5.setCellValue("计价数");
        cell_3_5.setCellStyle(styleMap.get("kaiti_center_12_style"));
        HSSFCell cell_3_6 = row_3.createCell(6);
        cell_3_6.setCellValue("单价");
        sheet.addMergedRegion(new CellRangeAddress(currentRowIndex + 3, currentRowIndex + 3,6,7));
        cell_3_6.setCellStyle(styleMap.get("kaiti_center_12_style"));
        HSSFCell cell_3_8 = row_3.createCell(8);
        cell_3_8.setCellValue("金额");
        cell_3_8.setCellStyle(styleMap.get("kaiti_center_12_style"));
        sheet.addMergedRegion(new CellRangeAddress(currentRowIndex + 3, currentRowIndex + 3,8,12));
        HSSFCell cell_3_13 = row_3.createCell(13);
        cell_3_13.setCellValue("备注");
        cell_3_13.setCellStyle(styleMap.get("kaiti_center_12_style"));
        sheet.addMergedRegion(new CellRangeAddress(currentRowIndex + 3, currentRowIndex + 3,13,14));

        // 回拨单合计
        double total = 0;
        for (int i = 0;i < 5;i++) {
            HSSFRow row = sheet.createRow(currentRowIndex + 4 + i);
            row.setHeightInPoints(lineHeight);
            sheet.addMergedRegion(new CellRangeAddress(currentRowIndex + 4 + i, currentRowIndex + 4 + i,1,2));
            sheet.addMergedRegion(new CellRangeAddress(currentRowIndex + 4 + i, currentRowIndex + 4 + i,6,7));
            sheet.addMergedRegion(new CellRangeAddress(currentRowIndex + 4 + i, currentRowIndex + 4 + i,8,12));
            sheet.addMergedRegion(new CellRangeAddress(currentRowIndex + 4 + i, currentRowIndex + 4 + i,13,14));
            int size = items.size();
            Item item = null;
            if (i < size) {
                item = items.get(i);
            }
            if (item != null) {
                HSSFCell cell_1 = row.createCell(1); //产品名称
                cell_1.setCellValue(item.getName());
                cell_1.setCellStyle(styleMap.get("style_SimSun_center_12"));
                HSSFCell cell_3 = row.createCell(3); // 单位
                cell_3.setCellValue(item.getUnit());
                cell_3.setCellStyle(styleMap.get("style_SimSun_center_12"));
                HSSFCell cell_4 = row.createCell(4); // 总发数
                int totalCount = (int)Double.parseDouble(item.getCount());
                cell_4.setCellValue(totalCount);
                cell_4.setCellStyle(styleMap.get("style_SimSun_center_12"));
                HSSFCell cell_5 = row.createCell(5); // 计价数
                cell_5.setCellValue(totalCount);
                cell_5.setCellStyle(styleMap.get("style_SimSun_center_12"));
                HSSFCell cell_6 = row.createCell(6); // 单价
                cell_6.setCellValue(Double.parseDouble(item.getPrice()));
                cell_6.setCellStyle(styleMap.get("moneyStyle"));
                HSSFCell cell_8 = row.createCell(8); // 金额
                cell_8.setCellStyle(styleMap.get("moneyStyle"));
                double price = totalCount * Double.parseDouble(item.getPrice());
                cell_8.setCellValue(price);
                HSSFCell cell_13 = row.createCell(13); // 备注
                total += price;
            } else {
                HSSFCell cell_8 = row.createCell(8); // 金额
                cell_8.setCellStyle(styleMap.get("moneyStyle"));
                cell_8.setCellValue(Double.parseDouble("0"));
            }

        }
        HSSFRow row_9 = sheet.createRow(currentRowIndex + 9);
        row_9.setHeightInPoints(lineHeight);
        HSSFCell cell_9_1 = row_9.createCell(1);
        cell_9_1.setCellValue("合计");
        cell_9_1.setCellStyle(styleMap.get("kaiti_center_12_style"));
        sheet.addMergedRegion(new CellRangeAddress(currentRowIndex + 9, currentRowIndex + 9,1,2));
        HSSFCell cell_9_6 = row_9.createCell(6);
        sheet.addMergedRegion(new CellRangeAddress(currentRowIndex + 9, currentRowIndex + 9,6,7));
        HSSFCell cell_9_8 = row_9.createCell(8);
        cell_9_8.setCellValue(total);
        HSSFCellStyle coin_style = wb.createCellStyle();
        HSSFDataFormat coinFormat = wb.createDataFormat();
        coin_style.setDataFormat(coinFormat.getFormat("¥#,##0.00"));
        HSSFFont font_SSimSun_center_12 = wb.createFont();
        font_SSimSun_center_12.setFontName("宋体"); //字体
        font_SSimSun_center_12.setFontHeightInPoints((short)12); //字体大小
        coin_style.setFont(font_SSimSun_center_12);
        coin_style.setAlignment(HorizontalAlignment.CENTER); // 水平居中
        cell_9_8.setCellStyle(coin_style);
        sheet.addMergedRegion(new CellRangeAddress(currentRowIndex + 9, currentRowIndex + 9,8,14));

        HSSFRow row_10 = sheet.createRow(currentRowIndex + 10);
        row_10.setHeightInPoints(lineHeight);
        HSSFCell cell_10_1 = row_10.createCell(1);
        cell_10_1.setCellValue("金额合计(大写)");
        cell_10_1.setCellStyle(styleMap.get("kaiti_center_12_style"));
        HSSFCell cell_10_3 = row_10.createCell(3);
        cell_10_3.setCellValue(new Rmb(total).toHanStr());
        cell_10_3.setCellStyle(styleMap.get("style_SimSun_center_12"));
        sheet.addMergedRegion(new CellRangeAddress(currentRowIndex + 10, currentRowIndex + 10,1,2));
        sheet.addMergedRegion(new CellRangeAddress(currentRowIndex + 10, currentRowIndex + 10,3,14));

        HSSFRow row_11 = sheet.createRow(currentRowIndex + 11);
        row_11.setHeightInPoints(lineHeight);
        HSSFCell cell_11_1 = row_11.createCell(1);
        cell_11_1.setCellValue("业务主管：");
        cell_11_1.setCellStyle(styleMap.get("style_font_KaiTi_bold_12"));
        HSSFCell cell_11_3 = row_11.createCell(3);
        cell_11_3.setCellValue("发货：");
        cell_11_3.setCellStyle(styleMap.get("style_font_KaiTi_bold_12"));
        HSSFCell cell_11_4 = row_11.createCell(4);
        cell_11_4.setCellValue("陈财友");
        cell_11_4.setCellStyle(styleMap.get("kaiti_style"));
        HSSFCell cell_11_6 = row_11.createCell(6);
        cell_11_6.setCellValue("收货：");
        cell_11_6.setCellStyle(styleMap.get("style_font_KaiTi_bold_12"));
        HSSFCell cell_11_7 = row_11.createCell(7);
        sheet.addMergedRegion(new CellRangeAddress(currentRowIndex + 11, currentRowIndex + 11,7,8));
        HSSFCell cell_11_9 = row_11.createCell(9);
        cell_11_9.setCellValue("承运：");
        cell_11_9.setCellStyle(styleMap.get("style_font_KaiTi_bold_12"));
        sheet.addMergedRegion(new CellRangeAddress(currentRowIndex + 11, currentRowIndex + 11,9,10));
        HSSFCell cell_11_11 = row_11.createCell(11);
        cell_11_11.setCellValue("制单：");
        cell_11_11.setCellStyle(styleMap.get("style_font_KaiTi_bold_12"));
        sheet.addMergedRegion(new CellRangeAddress(currentRowIndex + 11, currentRowIndex + 11,11,12));
        HSSFCell cell_11_13 = row_11.createCell(13);
        cell_11_13.setCellValue("马峰");
        cell_11_13.setCellStyle(styleMap.get("kaiti_style"));
        sheet.addMergedRegion(new CellRangeAddress(currentRowIndex + 11, currentRowIndex + 11,13,14));
        currentRowIndex +=13;
    }

    private void initSheet() {
        //创建HSSFSheet对象
        sheet = wb.createSheet("外调拨单");
        sheet.setColumnWidth(0, 750);
        sheet.setColumnWidth(1, 2850);
        sheet.setColumnWidth(2, 2930);
        sheet.setColumnWidth(3, 1440);
        sheet.setColumnWidth(4, 1990);
        sheet.setColumnWidth(5, 2180);
        sheet.setColumnWidth(6, 1600);
        sheet.setColumnWidth(7, 1280);
        sheet.setColumnWidth(8, 1880);
        sheet.setColumnWidth(9, (int)(256*2.8+250));
        sheet.setColumnWidth(10, (int)(256*2.8+250));
        sheet.setColumnWidth(11, (int)(256*1.9+250));
        sheet.setColumnWidth(12, (int)(256*2+250));
        sheet.setColumnWidth(13, (int)(256*5.8+400));
        sheet.setColumnWidth(14, 2500);
        sheet.setColumnWidth(15, 750);
    }

    private void initStyle() {
        HSSFCellStyle common_style = wb.createCellStyle();
        HSSFFont font_SimSun_12 = wb.createFont();
        font_SimSun_12.setFontName("宋体"); //字体
        font_SimSun_12.setFontHeightInPoints((short)12); //字体大小
        common_style.setVerticalAlignment(VerticalAlignment.CENTER); // 垂直居中
        common_style.setFont(font_SimSun_12);
        styleMap.put("common_style", common_style);

        HSSFCellStyle style_SimSun_center_12 = wb.createCellStyle();
        HSSFFont font_SSimSun_center_12 = wb.createFont();
        font_SSimSun_center_12.setFontName("宋体"); //字体
        font_SSimSun_center_12.setFontHeightInPoints((short)12); //字体大小
        style_SimSun_center_12.setAlignment(HorizontalAlignment.CENTER); // 水平居中
        style_SimSun_center_12.setVerticalAlignment(VerticalAlignment.CENTER); // 垂直居中
        style_SimSun_center_12.setFont(font_SSimSun_center_12);
        styleMap.put("style_SimSun_center_12", style_SimSun_center_12);

        HSSFCellStyle kaiti_style = wb.createCellStyle();
        HSSFFont font_kaiti_12 = wb.createFont();
        font_kaiti_12.setFontName("楷体_GB2312"); //字体
        font_kaiti_12.setFontHeightInPoints((short)12); //字体大小
        kaiti_style.setVerticalAlignment(VerticalAlignment.CENTER); // 垂直居中
        kaiti_style.setFont(font_kaiti_12);
        styleMap.put("kaiti_style", kaiti_style);

        HSSFCellStyle kaiti_center_12_style = wb.createCellStyle();
        kaiti_center_12_style.setFont(font_kaiti_12);
        kaiti_center_12_style.setAlignment(HorizontalAlignment.CENTER); // 水平居中
        kaiti_center_12_style.setVerticalAlignment(VerticalAlignment.CENTER); // 垂直居中
        styleMap.put("kaiti_center_12_style", kaiti_center_12_style);

        HSSFCellStyle style_font_KaiTi_bold_12 = wb.createCellStyle();
        HSSFFont font_simSun_bold_12 = wb.createFont();
        font_simSun_bold_12.setFontName("楷体_GB2312"); //字体
        font_simSun_bold_12.setFontHeightInPoints((short)12); //字体大小
        font_simSun_bold_12.setBold(true); //字体加粗
        style_font_KaiTi_bold_12.setFont(font_simSun_bold_12);
        style_font_KaiTi_bold_12.setVerticalAlignment(VerticalAlignment.CENTER); // 垂直居中
        styleMap.put("style_font_KaiTi_bold_12", style_font_KaiTi_bold_12);

        HSSFCellStyle style_SimSun_10 = wb.createCellStyle();
        style_SimSun_10.setRotation((short)255); // 文字竖列排版
        style_SimSun_10.setAlignment(HorizontalAlignment.CENTER); // 水平居中
        style_SimSun_10.setVerticalAlignment(VerticalAlignment.CENTER); // 垂直居中
        HSSFFont font_SimSun_10 = wb.createFont();
        font_SimSun_10.setFontName("宋体"); //字体
        font_SimSun_10.setFontHeightInPoints((short)10); //字体大小
        style_SimSun_10.setFont(font_SimSun_10);
        styleMap.put("style_SimSun_10", style_SimSun_10);

        HSSFCellStyle moneyStyle = wb.createCellStyle();
        moneyStyle.setAlignment(HorizontalAlignment.CENTER); // 水平居中
        moneyStyle.setFont(font_SimSun_12);
        moneyStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("#,##0.00"));
        styleMap.put("moneyStyle", moneyStyle);
    }
}
