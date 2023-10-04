package com.assignment.timecard;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class AssignmentTimecard {

	public static void main(String[] args) {

		Map<String, String> longShiftEmployees = getLongShiftEmployees();
		Map<String, String> breakEmployees = getBreakEmployees();

		try (FileOutputStream fos = new FileOutputStream(new File("output.txt"))) {
			String newline = System.getProperty("line.separator");
			StringBuilder sb = new StringBuilder();

			sb.append("Employees who have less than 10 hours of time between shifts but greater than 1 hour:")
					.append(newline)
					.append("---------------------------------------------------------------------------------")
					.append(newline);

			for (Map.Entry<String, String> entry : breakEmployees.entrySet()) {
				sb.append("Name: ").append(entry.getKey()).append("     Position: ").append(entry.getValue())
						.append(newline);
			}

			sb.append(newline);

			sb.append("Employees who have worked for more than 14 hours in a single shift:").append(newline)
					.append("---------------------------------------------------------------------------------")
					.append(newline);

			for (Map.Entry<String, String> entry : longShiftEmployees.entrySet()) {
				sb.append("Name: ").append(entry.getKey()).append("     Position: ").append(entry.getValue())
						.append(newline);
			}

			fos.write(sb.toString().getBytes());
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	// (b) Method to get the employees who have less than 10 hours of time between shifts but greater than 1 hour
	public static Map<String, String> getBreakEmployees() {
		Map<String, String> breakEmployees = new HashMap<>();
		Sheet sheet = getSheet();

		// Variables to track the value
		String previousEmployee = null;
		Date previousShiftEndTime = null;

		// Iterating through the sheet
		for (int rowNum = 1; rowNum <= 1483; rowNum++) {
			Row row = sheet.getRow(rowNum);
			if (row != null) {
				String name = row.getCell(7).getStringCellValue(); // Employee Name
				String position = row.getCell(1).getStringCellValue(); // Position status
				
				Cell startTime = row.getCell(2);
				Cell endTime = row.getCell(3);
				
				Date currentShiftStartTime = null; // Current shift in time
				Date currentShiftEndTime = null; // Current shift out time
				
				if (isCellNotBlank(startTime) && isCellNotBlank(endTime)) {
					double currentDateStartTime = Double.parseDouble(String.valueOf(startTime.getNumericCellValue()));
					double currentDateEndTime = Double.parseDouble(String.valueOf(endTime.getNumericCellValue()));

					// Excel's epoch for Windows (January 1, 1900)
					long excelEpochMillis = dateToMillis(1900, 1, 1);

					// MilliSeconds since epoch
					long startTimeMillisSinceEpoch = (long) ((currentDateStartTime - 1) * 24 * 60 * 60 * 1000)
							+ excelEpochMillis;
					long endTimeMillisSinceEpoch = (long) ((currentDateEndTime - 1) * 24 * 60 * 60 * 1000)
							+ excelEpochMillis;

					// Creating the date object from the calculated milliseconds
					currentShiftStartTime = new Date(startTimeMillisSinceEpoch);
					currentShiftEndTime = new Date(endTimeMillisSinceEpoch);

					// If both the date object are not null and the name of the employee is same as the previous employee then only calculate the difference

					if ((previousShiftEndTime != null && currentShiftStartTime != null)
							&& name.equals(previousEmployee)) {
						long timeDiffMillis = currentShiftEndTime.getTime() - currentShiftStartTime.getTime();
						long hours = timeDiffMillis / (60 * 60 * 1000);

						if ((hours >= 1 && hours <= 10) && !breakEmployees.containsKey(name)) {
							breakEmployees.put(name, position);
						}
					}
				}
				previousEmployee = name;
				previousShiftEndTime = currentShiftEndTime;
			}
		}

		return breakEmployees;
	}

	// (c) Method to get all the employees who has worked for more than 14hrs in a single shift
	public static Map<String, String> getLongShiftEmployees() {

		Sheet sheet = getSheet();
		Map<String, String> longShiftEmployees = new HashMap<>();

		// Iterating through the sheet
		for (int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
			Row row = sheet.getRow(rowNum);
			if (row != null) {

				String name = row.getCell(7).getStringCellValue(); // Employee Name
				String position = row.getCell(1).getStringCellValue(); // Position
				
				Cell cell = row.getCell(4); // Time Card Hours

				if (!cell.getStringCellValue().isEmpty()) {
					// Converting the cell value from String to double type
					double hoursWorked = Double.parseDouble(cell.getStringCellValue().replace(':', '.'));

					if (hoursWorked >= 14 && !longShiftEmployees.containsKey(name)) {
						longShiftEmployees.put(name, position);
					}
				}
			}
		}
		return longShiftEmployees;
	}

	// Getting the Sheet
	private static Sheet getSheet() {
		try {
			String filePath = "D:\\Downloads\\Assignment_Timecard.xlsx";
			FileInputStream fis = new FileInputStream(filePath);

			Workbook workbook = new XSSFWorkbook(fis);
			Sheet sheet = workbook.getSheet("Sheet1");
			return sheet;
		} catch (IOException e) {
			e.printStackTrace();
		}
		return null;
	}

	// Checking if the cell is blank or not
	public static boolean isCellNotBlank(Cell cell) {
		boolean b = false;
		if (cell.getCellType() == CellType.STRING) {
			b = !cell.getStringCellValue().isEmpty();

		} else if (cell.getCellType() == CellType.NUMERIC) {
			b = !String.valueOf(cell.getNumericCellValue()).isEmpty();
		}
		return b;
	}

	// method to convert the date to milliseconds since epoch
	private static long dateToMillis(int year, int month, int day) {
		java.util.Calendar calendar = Calendar.getInstance();
		calendar.set(year, month - 1, day, 0, 0, 0);
		calendar.set(Calendar.MILLISECOND, 0);
		return calendar.getTimeInMillis();
	}

}
