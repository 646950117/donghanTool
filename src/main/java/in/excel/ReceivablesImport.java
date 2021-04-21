package in.excel;

import in.excel.po.DebtItem;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

public class ReceivablesImport implements ExcelImport {

    private String path;

    private Sheet inSheet;

    private int totalRows;

    // 标题行号
    private int titleRowNum;

    // 数据行号
    private int dataRowNum;

    // 物品名称
    List<String> itemName = new ArrayList<>();

    // 物品单位
    List<String> itemUnit = new ArrayList<>();

    public ReceivablesImport(String path) throws Exception {
        File dataFile = new File(path);
        File[] files = dataFile.listFiles();
        if (files.length > 1) {
            System.out.println("数据文件只能有一个，删除data下多余的excel！");
            return;
        }
        if (files[0].isDirectory()) {
            System.out.println("数据文件不能是目录，应放一个excel文件！");
            return;
        }

        File excelFile = files[0];
        if (!excelFile.exists()) {
            System.out.println("导入的数据文件不存在，路径：[" + path + "]");
            return;
        }
        System.out.println("开始加载数据文件：" + excelFile.getAbsolutePath());
        Workbook workbook = WorkbookFactory.create(new FileInputStream(excelFile));
        int sheetSize = workbook.getNumberOfSheets();
        if (sheetSize == 0) {
            System.out.println("数据文件中没有数据！");
            return;
        }
        int dataIndex = -1;
        for (int i = 0;i < sheetSize;i++) {
            Sheet tempSheet = workbook.getSheetAt(i);
            int tempTotalRows = tempSheet.getLastRowNum() + 1;
            if (tempTotalRows <= 3) {
                continue;
            }
            boolean flag = false;
            for (int j = 0;j < tempTotalRows;j++) {
                Row tempRow = tempSheet.getRow(j + 1);
                Cell tempCell = tempRow.getCell(0);
                if (tempCell != null) {
                    if ("代码".equals(tempCell.toString())) {
                        flag = true;
                        break;
                    }
                }
            }
            if (flag) {
                dataIndex = i;
                break;
            }
        }

        if (dataIndex == -1) {
            System.out.println("数据文件中没有符合的数据！");
            return;
        }
        System.out.println("从数据文件中第" + (dataIndex + 1) + "个sheet中获取数据！");
        inSheet = workbook.getSheetAt(dataIndex);
        totalRows = inSheet.getLastRowNum() + 1;
        for (int i = 0;i < totalRows;i++) {
            Row row = inSheet.getRow(i + 1);
            Cell codeCell = row.getCell(0);
            Cell customerCell = row.getCell(1);

            if (codeCell == null || customerCell == null) {
                continue;
            }
            if ("代码".equals(codeCell.toString()) && "客户名称".equals(customerCell.toString())) {
                titleRowNum = i + 1;
                dataRowNum = i + 3;
                Row unitRow = inSheet.getRow(i + 2);
                for (int j = 0;;j++) {
                    Cell countCell = unitRow.getCell(j * 3 + 2);
                    Cell unitCell = unitRow.getCell(j * 3 + 3);
                    Cell priceCell = unitRow.getCell(j * 3 + 4);
                    if (countCell == null || unitCell == null || priceCell == null) {
                        break;
                    }
                    if ("数量".equals(countCell.toString()) && "单价".equals(unitCell.toString())
                            && "金额".equals(priceCell.toString())) {
                        itemName.add(row.getCell(j * 3 + 2).toString());
                        String commentText = row.getCell(j * 3 + 2).getCellComment().getString().getString();
                        commentText = commentText.split("\n")[1];
                        itemUnit.add(commentText);
                    } else {
                        break;
                    }
                }
                break;
            }

            // 前10行没有符合的数据，则抛错
            if (i > 10) {
                System.out.println("导入的数据表格格式不对！");
                return;
            }
        }
    }


    @Override
    public List<DebtItem> loadExcel() {
        List<DebtItem> result = new ArrayList<>();
        for (int i = dataRowNum;i < totalRows;i++) {
            Row row = inSheet.getRow(i);
            Cell firstCell = row.getCell(0);
            if (firstCell == null) {
                break;
            }
            DebtItem debtItem = build(row);
            result.add(debtItem);
        }
        System.out.println("一共加载到[" + result.size() + "]条数据！");
        return result;
    }

    private DebtItem build(Row row) {
        String code = row.getCell(0).toString();
        String customer = row.getCell(1).toString();
        DebtItem debtItem =  new DebtItem();
        debtItem.setCode(code);
        debtItem.setCustom(customer);
        for (int i = 0;i < itemName.size();i++) {
            String name = itemName.get(i);
            String unit = itemUnit.get(i);
            Cell countCell = row.getCell(i * 3 + 2); //数量
            countCell.setCellType(CellType.STRING);
            Cell unitCell = row.getCell(i * 3 + 3); //单价
            Cell priceCell = row.getCell(i * 3 + 4); //金额
            if (countCell != null && !"".equals(countCell.toString()) && !"0".equals(countCell.toString())) {
                if (unitCell == null || priceCell == null) {
                    throw new IllegalStateException("代码：[" + code + "], 物品：[" + name + "]数据有误，请检查后再执行！");
                }
                debtItem.addItem(name, unit, String.valueOf(countCell.getNumericCellValue()), String.valueOf(unitCell.getNumericCellValue()), priceCell.toString());
            }
        }
        return debtItem;
    }
}
