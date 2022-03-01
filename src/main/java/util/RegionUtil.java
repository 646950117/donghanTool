package util;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.util.HSSFCellUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

public class RegionUtil {

    public static void main(String[] args) {
        String customerId = "LRCI2017031400019";
        System.out.println(conversion(customerId));
    }

    public static String conversion(String code){
        char[] chars = code.toCharArray();
        for(int i = 0 ; i < chars.length ; i++){
            char ch = chars[i];
            if(ch >= 97 && ch <= 122){
                //do nothing
            }else if(ch >= 65 && ch <= 90){
                chars[i] = (char)(ch + 32);
            }else if(ch >= 48 && ch <= 57){
                //do nothing
            }
            else{
                chars[i]= '#';
            }
        }
        return new String(chars).replace("#","");
    }

    public static  void setRegionStyle(HSSFSheet sheet, CellRangeAddress region, HSSFCellStyle cs){
        for(int i=region.getFirstRow();i<=region.getLastRow();i++) {
            HSSFRow row= HSSFCellUtil.getRow(i,sheet);
                for(int j=region.getFirstColumn();j<=region.getLastColumn();j++){
                    HSSFCell cell=HSSFCellUtil.getCell(row,(short)j);
                    if (cell == null) {
                        cell = row.createCell(j);
                        cell.setCellValue("");
                    }
                    cell.setCellStyle(cs);
                }

        }

    }

    private static void setBorder(HSSFCellStyle style) {
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
    }
}
