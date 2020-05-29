package util;

import model.Gender;
import model.User;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.util.LinkedList;
import java.util.List;

public class XLSXUtil {

    public static List<User> reedUsers(String path) {
        List<User> result = new LinkedList<>();
        try {
            XSSFWorkbook workbook = new XSSFWorkbook(path);
            Sheet sheet = workbook.getSheetAt(0);
            int lastRowNum = sheet.getLastRowNum();
            for (int i = 1; i <= lastRowNum; i++) {
                Row row = sheet.getRow(i);
                  result.add(getUserFromRow(row));
            }
        } catch (IOException e) {
            e.printStackTrace();
            System.err.println("Error while importing users");
        }
        return result;
    }

    private static User getUserFromRow(Row row) {
        Cell phoneNumber = row.getCell(4);
        String phoneNumberStr = phoneNumber.getCellType() == CellType.NUMERIC ?
                String.valueOf(Double.valueOf(phoneNumber.getNumericCellValue()).intValue()) : phoneNumber.getStringCellValue();
        Cell password = row.getCell(5);
        String passwordStr = password.getCellType() == CellType.NUMERIC ?
                String.valueOf(Double.valueOf(password.getNumericCellValue()).intValue()) : password.getStringCellValue();
        return User.builder()
                .name(row.getCell(0).getStringCellValue())
                .surname(row.getCell(1).getStringCellValue())
                .age(Double.valueOf(row.getCell(2).getNumericCellValue()).intValue())
                .gender(Gender.valueOf(row.getCell(3).getStringCellValue()))
                .phoneNumber(phoneNumberStr)
                .password(passwordStr)
                .build();
    }
}
