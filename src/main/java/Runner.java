
import in.excel.ExcelImport;
import in.excel.ReceivablesImport;
import in.excel.po.DebtItem;
import template.DeptExcelTemplate;
import template.ExcelTemplate;

import java.util.List;

public class Runner {
    public static void main(String[] args) throws Exception {
        ExcelImport<DebtItem> ipt = new ReceivablesImport("data");
        List<DebtItem> debts = ipt.loadExcel();
        ExcelTemplate template = new DeptExcelTemplate("output", debts, args[0], args[1], args[2]);
        template.write();
    }
}
