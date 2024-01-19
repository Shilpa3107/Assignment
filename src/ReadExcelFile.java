package src;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelFile {
    public static void main(String[] args) {
        try {
            // Loading Excel file
            FileInputStream fileInputStream = new FileInputStream(new File("Assignment_Timecard.xlsx"));
            Workbook workbook = new XSSFWorkbook(fileInputStream);
            Sheet sheet = workbook.getSheetAt(0);

            // Store employee data
            Map<String, List<Shift>> employeeShifts = new HashMap<>();

            // Read data from Excel
            Iterator<Row> rowIterator = sheet.iterator();
            rowIterator.next(); // Skip header row
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                String employeeName = row.getCell(7).getStringCellValue();
                Cell timeInCell = row.getCell(2);
                Cell timeOutCell = row.getCell(3);
                
                Date timeIn, timeOut;
                
                try {
                    if (timeInCell == null || timeOutCell == null ||
                    timeInCell.getCellType() == CellType.BLANK || timeOutCell.getCellType() == CellType.BLANK) {
                    // Skip empty cells
                    continue;
                }
                    if (timeInCell.getCellType() == CellType.NUMERIC || timeInCell.getCellType() == CellType.FORMULA) {
                        timeIn = timeInCell.getDateCellValue();
                    } else if (timeInCell.getCellType() == CellType.STRING) {
                        timeIn = new SimpleDateFormat("MM/dd/yyyy hh:mm a").parse(timeInCell.getStringCellValue());
                    } else {
                        // Handle other cell types if necessary
                        continue;
                    }
                    
                    if (timeOutCell.getCellType() == CellType.NUMERIC || timeOutCell.getCellType() == CellType.FORMULA) {
                        timeOut = timeOutCell.getDateCellValue();
                    } else if (timeOutCell.getCellType() == CellType.STRING) {
                        timeOut = new SimpleDateFormat("MM/dd/yyyy hh:mm a").parse(timeOutCell.getStringCellValue());
                    } else {
                        // Handle other cell types if necessary
                        continue;
                    }
                
                    // Calculate shift duration
                    long shiftDuration = (timeOut.getTime() - timeIn.getTime()) / (60 * 1000);
                
                    // Create Shift object
                    Shift shift = new Shift(timeIn, timeOut, shiftDuration);
                
                    // Update employee shifts
                    employeeShifts.computeIfAbsent(employeeName, k -> new ArrayList<>()).add(shift);
                
                } catch (ParseException | IllegalStateException e) {
                    System.out.println("Error parsing date. Skipping row.");
                    // Optionally continue to the next iteration to avoid further issues with this row
                    e.printStackTrace();
                    continue;
                }
            }
            

            // Analyze and print results
            analyzeAndPrintResults(employeeShifts);

            // Close the workbook
            workbook.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void analyzeAndPrintResults(Map<String, List<Shift>> employeeShifts) {
        for (Map.Entry<String, List<Shift>> entry : employeeShifts.entrySet()) {
            String employeeName = entry.getKey();
            List<Shift> shifts = entry.getValue();

            System.out.println("Employee: " + employeeName);

            // Identify employees who worked for 7 consecutive days
            if (hasConsecutiveDays(shifts, 7)) {
                System.out.println("Worked for 7 consecutive days");
            }

            //Identify employees who had less than 10 hours between shifts but greater than 1 hour
            if (hasShortBreaks(shifts, 10, 1)) {
                System.out.println("Had less than 10 hours between shifts but greater than 1 hour");
            }

            // Identify employees who worked for more than 14 hours in a single shift
            if (hasLongShifts(shifts, 14 * 60)) { // 14 hours in minutes
                System.out.println("Worked for more than 14 hours in a single shift");
            }

            System.out.println("-----------------------------");
        }
    }

    private static boolean hasConsecutiveDays(List<Shift> shifts, int consecutiveDays) {
        // Sort shifts by time in
        shifts.sort(Comparator.comparing(Shift::getTimeIn));

        // Check for consecutive days
        int consecutiveCount = 1;
        for (int i = 1; i < shifts.size(); i++) {
            long dayDifference = daysBetween(shifts.get(i - 1).getTimeIn(), shifts.get(i).getTimeIn());
            if (dayDifference == 1) {
                consecutiveCount++;
                if (consecutiveCount == consecutiveDays) {
                    return true;
                }
            } else {
                consecutiveCount = 1; // Reset count if not consecutive
            }
        }

        return false;
    }

    private static boolean hasShortBreaks(List<Shift> shifts, int maxBreak, int minBreak) {
        // Sort shifts by time in
        shifts.sort(Comparator.comparing(Shift::getTimeIn));

        // Check for short breaks
        for (int i = 1; i < shifts.size(); i++) {
            long breakDuration = (shifts.get(i).getTimeIn().getTime() - shifts.get(i - 1).getTimeOut().getTime()) / (60 * 1000);

            if (breakDuration < maxBreak && breakDuration > minBreak) {
                return true;
            }
        }

        return false;
    }

    private static boolean hasLongShifts(List<Shift> shifts, int threshold) {
        for (Shift shift : shifts) {
            if (shift.getDuration() > threshold) {
                return true;
            }
        }
        return false;
    }

    private static long daysBetween(Date date1, Date date2) {
        long diff = Math.abs(date1.getTime() - date2.getTime());
        return diff / (24 * 60 * 60 * 1000);
    }
}

class Shift {
    private Date timeIn;
    private Date timeOut;
    private long duration; // in minutes

    public Shift(Date timeIn, Date timeOut, long duration) {
        this.timeIn = timeIn;
        this.timeOut = timeOut;
        this.duration = duration;
    }

    public Date getTimeIn() {
        return timeIn;
    }

    public Date getTimeOut() {
        return timeOut;
    }

    public long getDuration() {
        return duration;
    }
}
