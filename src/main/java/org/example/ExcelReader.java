package org.example;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {

  public static final String fileExcelPath = "src/main/resources/BangCong.xlsx";

  public static void main(String[] args) {

    List<Employee> employees = new ArrayList<>();
    try (InputStream fis = new FileInputStream(new File(fileExcelPath));
        XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

      XSSFSheet sheet = workbook.getSheetAt(0);
//      Iterator<Row> rowIterator = sheet.iterator();
      Map<String, Employee> employeeMap = new HashMap<>();
      List<String> shifts = new LinkedList<>();
      int shiftIndex = 3;
      Row shiftRow = sheet.getRow(5);
      Cell shiftCell = shiftRow.getCell(shiftIndex);
      while (shiftCell != null && shiftCell.getCellType() != CellType.BLANK) {
        if (shiftCell.getStringCellValue().equals("$")) {
          shiftCell = shiftRow.getCell(++shiftIndex);
          continue;
        }
        shifts.add(shiftCell.getStringCellValue());
        shiftCell = shiftRow.getCell(++shiftIndex);
      }
      List<LinkedList<HashMap<String, Double>>> dayList = new LinkedList<>();

      LinkedList<HashMap<String, Double>> mList = new LinkedList<>();

      for (Row row : sheet) {
        Map<String, Double> dayData = new HashMap<>();
        int dayIndex = 17;
        if (row.getRowNum() < 6) {
          continue;
        }
        if (row.getCell(0) == null) {
          break;
        }
        Row titleRow = sheet.getRow(5);
        Row dayRowCheck = sheet.getRow(3);
        for (Cell cell : row) {
          int colIndex = cell.getColumnIndex();

          // Skip if the cell is in the first 4 columns is not a number
          if (colIndex <= shiftIndex || cell.getCellType() == CellType.STRING
              || cell.getCellType() == CellType.FORMULA) {
            continue;
          }


          Cell titleCell = titleRow.getCell(colIndex);
          String dayTitle = titleCell.getStringCellValue();

          if (cell.getCellType() == CellType.NUMERIC ) {

            if (dayRowCheck.getCell(dayIndex).getCellType() != CellType.BLANK) {
              mList.add(new HashMap<>(dayData));
              dayData.clear();

            }
            double dataValue = cell.getNumericCellValue();
            dayData.put(dayTitle, dataValue);
            dayIndex++;
          } else if (cell.getCellType() == CellType.BLANK) {
            if (dayRowCheck.getCell(dayIndex).getCellType() != CellType.BLANK && !dayData.isEmpty()) {
              mList.add(new HashMap<>(dayData));
              dayData.clear();
            }
            dayData.put(dayTitle, 0.0);
            dayIndex++;
          }
        }
        if (!dayData.isEmpty()) {
          mList.add(new HashMap<>(dayData) );
        }
        dayList.add(new LinkedList<>(mList));
        mList.clear();
      }

      LinkedList<Employee> employeeList = new LinkedList<>();
      for (Row row : sheet) {
        if (row.getRowNum() < 6) continue;

        if (row.getCell(0).getCellType() == CellType.BLANK) break;

        Employee employee = new Employee();
        Cell idCell = row.getCell(1);
        Cell nameCell = row.getCell(2);

        if (idCell != null && idCell.getCellType() == CellType.STRING) {
          employee.setCode(idCell.getStringCellValue());
        }
        if (nameCell != null && nameCell.getCellType() == CellType.STRING) {
          employee.setName(nameCell.getStringCellValue());
        }

        employeeList.add(employee);
      }

      LinkedList<HashMap<String, Double>> salaryTable = new LinkedList<>();
      for (Row row : sheet) {
        if (row.getRowNum() < 6) continue;
        if (row.getCell(0) == null || row.getCell(0).getCellType() == CellType.BLANK) break;

        HashMap<String, Double> salaryData = new HashMap<>();

        for (Cell cell : row) {
          int colIndex = cell.getColumnIndex();
          if (colIndex < 3) continue;

          Row titleRow = sheet.getRow(5);
          Cell titleCell = titleRow.getCell(colIndex);

          if (titleCell == null || titleCell.getCellType() == CellType.BLANK) continue;

          switch (cell.getCellType()) {
            case NUMERIC, FORMULA -> {
              if (titleCell.getStringCellValue().equals("$")) {
                processSalaryCell(sheet, colIndex, salaryData, cell.getNumericCellValue());
              }
            }
            case BLANK -> {
              if (titleCell.getStringCellValue().equals("$")) {
                processSalaryCell(sheet, colIndex, salaryData, 0.0);
              }
            }
          }
          if (colIndex >= shiftIndex) break;
        }
        salaryTable.add(new HashMap<>(salaryData));
      }

      LinkedList<Double> fileSalary = new LinkedList<>();

      for (Row row : sheet) {
        if (row.getRowNum() < 6) continue;


        if (row.getCell(0).getCellType() == CellType.BLANK) {
          break;
        }
        for (Cell cell : row){
          if(cell.getColumnIndex() != shiftIndex){
            continue;
          }
          fileSalary.add(cell.getNumericCellValue());
        }
      }

      System.out.println();
      showEmployee(employeeList, salaryTable, dayList, fileSalary);


    } catch (IOException e) {
      throw new RuntimeException(e);
    }

  }

  private static void showEmployee(LinkedList<Employee> employeeList, LinkedList<HashMap<String, Double>> salaryTable, List<LinkedList<HashMap<String, Double>>> dayList, LinkedList<Double> fileSalary) {
    int totalIndex = employeeList.size();
    for (int i = 0; i < totalIndex; i++) {
      Employee currentEmp = employeeList.get(i);
      System.out.printf("\tID            : %s%n", currentEmp.getCode());
      System.out.printf("\tEmployee name : %s%n", currentEmp.getName());

      LinkedList<HashMap<String, Double>> eachDay = dayList.get(i);
      HashMap<String, Double> eSalary = salaryTable.get(i);
      double monthSalary = 0;

      for (int j = 0; j < eachDay.size(); j++) {
        HashMap<String, Double> eDay = eachDay.get(j);
        double totalDayWorkHour = eDay.values().stream().mapToDouble(Double::doubleValue).sum();

        if (totalDayWorkHour == 0) continue;

        System.out.printf("\tDay %2d [ ", j + 1);
        double daySalary = 0;

        for (Map.Entry<String, Double> entry : eDay.entrySet()) {
          String workType = entry.getKey();
          double hours = entry.getValue();
          if (hours > 0) {
            daySalary += eSalary.get(workType) * hours;
            System.out.printf("%s: %.2f | ", workType, hours);
          }
        }

        System.out.printf("Total work hour: %.2f | Day salary: %.0f ]%n", totalDayWorkHour, daySalary);
        monthSalary += daySalary;
      }

      System.out.println("\tTotal salary : " + String.format("%.0f",monthSalary));
      if(String.format("%.0f",fileSalary.get(i)).equals(String.format("%.0f",monthSalary))){
        System.out.println("\tTotal salary is equals with data from file!");
      }
      else{
        System.out.println("\tTotal salary is not equals with data from file!");
      }
    }
  }


  private static void processSalaryCell(Sheet sheet, int colIndex,
      HashMap<String, Double> salaryData, double cellValue) {
    Row titleRow = sheet.getRow(5);
    int reverseIndex = 1;

    while (colIndex - reverseIndex > 0) {
      Cell titleCell = titleRow.getCell(colIndex - reverseIndex);
      if (titleCell != null && titleCell.getCellType() == CellType.STRING) {
        String title = titleCell.getStringCellValue();
        if (title.equals("$")) break;
        salaryData.put(title, cellValue);
      }
      reverseIndex++;
    }
  }
}
