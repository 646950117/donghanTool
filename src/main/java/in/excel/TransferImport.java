package in.excel;

import in.excel.po.TransferItem;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.util.*;

public class TransferImport implements ExcelImport{

    private String path;

    private static String reg = "\\d{1,2}-\\d{1,2}";

    private Workbook workbook;

    private Map<String, Map<String, String>> accountMap;

    public TransferImport(String inPath) throws Exception {
        this.path = inPath;
        this.accountMap = new HashMap<>();
        File dataFile = new File(path);
        File[] files = dataFile.listFiles();
        if (files.length != 2) {
            throw new IllegalStateException("data文件内文件个数不对，应为《客户转账信息.xls》和对账单表格");
        }
        File excelFile = null;
        File accountFile = null;
        for (int i = 0;i < files.length;i++) {
            File file = files[i];
            if ("客户转账信息.xls".equals(file.getName())) {
                accountFile = file;
            } else {
                excelFile = file;
            }
        }
        if (excelFile == null || !excelFile.exists()) {
            throw new IllegalStateException("data文件夹下不存在对账单表格！");
        }
        if (accountFile == null || !accountFile.exists()) {
            throw new IllegalStateException("data文件夹下不存在《客户转账信息.xls》！");
        }

        Workbook accountWorkbook = WorkbookFactory.create(new FileInputStream(accountFile));
        int sheetSize = accountWorkbook.getNumberOfSheets();
        if (sheetSize == 0) {
            System.out.println("没找到客户转账信息！");
            return;
        }
        Sheet accountSheet = accountWorkbook.getSheetAt(0);
        int rowNum = accountSheet.getLastRowNum() + 1;
        for (int i = 0;i < rowNum;i++) {
            Row row = accountSheet.getRow(i + 1);
            if (row != null) {
                Cell cell = row.getCell(2);
                if (cell == null) {
                    continue;
                }

                String account = cell.getStringCellValue();
                if (account == null || account.length() == 0) {
                    continue;
                }
                account = account.trim();
                Cell cardCell = row.getCell(3);
                String card = "";
                if (cardCell != null) {
                    card = cardCell.getStringCellValue();
                }
                Map<String, String> map = new HashMap();
                map.put("card", card);

                Cell bankCell = row.getCell(4);
                String bank = "";
                if (bankCell != null) {
                    bank = bankCell.getStringCellValue();
                }
                map.put("bank", bank);
                accountMap.put(account, map);
            }
        }
        if (accountMap.size() == 0) {
            throw new IllegalStateException("没有加载客户转账信息，检查《客户转账信息.xls》格式是否有误！");
        }

        workbook = WorkbookFactory.create(new FileInputStream(excelFile));
        sheetSize = workbook.getNumberOfSheets();
        if (sheetSize == 0) {
            System.out.println("数据文件中没有数据！");
            return;
        }
    }

    @Override
    public List<TransferItem> loadExcel() {
        List<TransferItem> result = new ArrayList<>();
        // 获取第一个sheet
        Sheet tempSheet = workbook.getSheetAt(0);
        int tempTotalRows = tempSheet.getLastRowNum() + 1;
        int noMatchNum = 0;
        int thresholdRowNum = 50;
        for (int i = 0;i < tempTotalRows;i++) {
            Row tempRow = tempSheet.getRow(i);
            if (tempRow == null) {
                if (++noMatchNum > thresholdRowNum) {
                    // 超过阈值，说明没有数据直接不往下面读了
                    break;
                }
            } else {
                noMatchNum = 0;
                Cell tempCell = tempRow.getCell(0);
                if (tempCell == null) {
                    continue;
                }
                String date = tempCell.getStringCellValue();
                if (!date.matches(reg)) {
                    if (++noMatchNum > thresholdRowNum) {
                        // 超过阈值，说明没有数据直接不往下面读了
                        break;
                    }
                } else {
                    noMatchNum = 0;
                    Cell AccountCell = tempRow.getCell(4);
                    HSSFColor fillColor = (HSSFColor)AccountCell.getCellStyle().getFillForegroundColorColor();
                    if ("FFFF:CCCC:0".equals(fillColor.getHexString())) {
                        TransferItem item = new TransferItem();
                        item.setMonth(date.split("-")[0]);
                        item.setDay(date.split("-")[1]);
                        item.setBusinessType(tempRow.getCell(1).getStringCellValue());
                        String account = tempRow.getCell(4).getStringCellValue().trim();
                        item.setPayAccount(account);
                        Map<String, String> accountInfo = accountMap.get(account);
                        if (accountInfo != null) {
                            item.setPayCard(accountInfo.get("card"));
                            item.setPayBank(accountInfo.get("bank"));
                        } else {
                            Collection collection = accountMap.values();
                            accountInfo = (Map<String, String>) new ArrayList<>(collection).get(new Random().nextInt(collection.size()));
//                            System.out.println("【警告】回单编号：[" + item.getTransferNum() + "]，客户：[" + account + "]，没有找到客户转账信息！");
                            item.setPayCard(accountInfo.get("card"));
                            item.setPayBank(accountInfo.get("bank"));
                        }
                        item.setPrice(tempRow.getCell(9).getNumericCellValue());
                        item.setTips(tempRow.getCell(7).getStringCellValue());
//                        item.setPayBank(tempRow.getCell(4).getStringCellValue());
                        result.add(item);
                    }


                }
            }

        }
        return result;
    }
}
