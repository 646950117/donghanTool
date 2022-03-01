package in.excel;

import java.io.FileNotFoundException;
import java.util.List;

public interface ExcelImport<T> {

    List<T> loadExcel() throws FileNotFoundException;

}
