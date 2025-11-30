package com.assignment.excelhandler;

import java.io.*;
import java.util.Scanner;
import org.apache.poi.xssf.usermodel.*;

public class ExcelFileHandler {

    private String filePath;
    private XSSFWorkbook workbook;
    private XSSFSheet sheet;

    public ExcelFileHandler(String filePath) {
        this.filePath = filePath;
        loadFile();
    }

    private void loadFile() {
        try (FileInputStream fis = new FileInputStream(filePath)) {
            workbook = new XSSFWorkbook(fis);
            sheet = workbook.getSheetAt(0);
            System.out.println("File loaded successfully!\n");
        } catch (Exception e) {
            System.out.println("Failed to load file: " + e.getMessage());
        }
    }

    public void displaySheet() {
        if (sheet == null) {
            System.out.println("Sheet not found.");
            return;
        }

        System.out.println("\n---- Excel Data ----");
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            XSSFRow row = sheet.getRow(i);
            if (row == null) continue;

            for (int j = 0; j < row.getLastCellNum(); j++) {
                XSSFCell cell = row.getCell(j);
                System.out.print((cell == null ? "" : cell.toString()) + "\t");
            }
            System.out.println();
        }
        System.out.println("---------------------\n");
    }

    public void updateCell(int rowIndex, int colIndex, String value) {
        if (sheet == null) return;

        XSSFRow row = sheet.getRow(rowIndex);
        if (row == null) row = sheet.createRow(rowIndex);

        XSSFCell cell = row.getCell(colIndex);
        if (cell == null) cell = row.createCell(colIndex);

        cell.setCellValue(value);
        System.out.println("Updated cell (" + rowIndex + "," + colIndex + ") successfully!\n");
    }

    public void save() {
        try (FileOutputStream fos = new FileOutputStream(filePath)) {
            workbook.write(fos);
            workbook.close();
            System.out.println("File saved successfully!");
        } catch (IOException e) {
            System.out.println("Error saving file: " + e.getMessage());
        }
    }

    public static void main(String[] args) {

        Scanner sc = new Scanner(System.in);

        ExcelFileHandler handler = new ExcelFileHandler("C:\\Users\\sachi\\Desktop\\Sample.xlsx");

        while (true) {
            System.out.println("\nChoose an option:");
            System.out.println("1. Display Excel Sheet");
            System.out.println("2. Update a Cell");
            System.out.println("3. Save & Exit");
            System.out.print("Enter choice: ");
            
            int choice = sc.nextInt();

            if (choice == 1) {
                handler.displaySheet();

            } else if (choice == 2) {
                System.out.print("Enter row index: ");
                int row = sc.nextInt();

                System.out.print("Enter column index: ");
                int col = sc.nextInt();

                sc.nextLine(); // clear buffer

                System.out.print("Enter new value: ");
                String value = sc.nextLine();

                handler.updateCell(row, col, value);

            } else if (choice == 3) {
                handler.save();
                break;

            } else {
                System.out.println("Invalid option. Try again.");
            }
        }

        sc.close();
    }
}
