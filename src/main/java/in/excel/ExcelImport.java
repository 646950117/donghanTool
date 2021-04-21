package in.excel;

import java.util.List;

public interface ExcelImport<T> {

    List<T> loadExcel();

}
