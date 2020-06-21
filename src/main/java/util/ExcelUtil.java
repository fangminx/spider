package util;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * excel工具
 */
public class ExcelUtil {
    private ExcelUtil() {
    }

    /**
     * 获得指定路径的excel工作簿文件
     *
     * @param fileCompletePath
     * @return
     */
    public static Workbook getWorkbook(String fileCompletePath) {
        Workbook result = null;
        InputStream is = null;
        String suffixName = fileCompletePath.substring(fileCompletePath.lastIndexOf(".") + 1);
        try {
            is = new FileInputStream(fileCompletePath);
            if (suffixName.equals("xls")) {
                result = new HSSFWorkbook(is);
            } else if (suffixName.equals("xlsx")) {
                result = new XSSFWorkbook(is);
            }

            is.close();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (is != null) {
                    is.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
            return result;
        }
    }

    /**
     * 以字符串形式获得单元格内的值，若单元格不存在，返回""
     *
     * @param cell
     * @return
     */
    public static String getCellVal(Cell cell) {
        if (cell == null) return "";

            if(cell.getCellType().equals(CellType.BOOLEAN)) {
                return Boolean.toString(cell.getBooleanCellValue());
            }

        return "";
    }
}
