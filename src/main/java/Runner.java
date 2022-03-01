
import constant.CommonConstant;
import in.excel.ExcelImport;
import in.excel.ReceivablesImport;
import in.excel.TransferImport;
import in.excel.po.DebtItem;
import in.excel.po.TransferItem;
import template.DeptExcelTemplate;
import template.ExcelTemplate;
import template.TransferExcelTemplate;

import java.util.Collections;
import java.util.List;

public class Runner {
    public static void main(String[] args) throws Exception {
        String orderType = args[2];
        if (orderType == null || orderType.length() ==0) {
            throw new IllegalStateException("回单类型为空！");
        }

        switch (orderType) {
            case CommonConstant.DEPT_ORDER:
                ExcelImport<DebtItem> ipt = new ReceivablesImport("data");
                List<DebtItem> debts = ipt.loadExcel();
                Collections.shuffle(debts);
                ExcelTemplate template = new DeptExcelTemplate("output", debts, args[0], args[1]);
                template.write();
                break;
            case CommonConstant.TRANSFER_ORDER:
                ExcelImport<TransferItem> eit = new TransferImport("data");
                List<TransferItem> transferItems = eit.loadExcel();
                template = new TransferExcelTemplate(transferItems, "output",args[0], args[1]);
                template.write();
                break;
        }

    }
}
