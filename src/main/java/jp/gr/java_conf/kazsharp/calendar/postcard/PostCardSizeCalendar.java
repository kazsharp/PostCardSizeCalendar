package jp.gr.java_conf.kazsharp.calendar.postcard;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.Year;
import java.time.YearMonth;
import java.util.stream.IntStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.ShapeTypes;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import com.samuraism.holidays.日本の祝休日;

public class PostCardSizeCalendar {

  private static final int TARGET_YEAR = 2021;

  private static final String OUTPUT_FILE = "C:\\cal\\calender2021.xlsx";

  private Workbook wb;

  private PostCardSizeCalendarStyle style;

  private 日本の祝休日 holidays = new 日本の祝休日();

  public static void main(String[] args) throws Exception {

    new PostCardSizeCalendar().createCalendar();
  }

  private void createCalendar() throws IOException {

    wb = WorkbookFactory.create(true);
    style = new PostCardSizeCalendarStyle(wb);

    Year year = Year.of(TARGET_YEAR);
    IntStream.rangeClosed(1, 12).forEach(i -> {
      YearMonth ym = year.atMonth(i);
      createSheet(ym);
      printSetting(i);
    });

    try (OutputStream out = new FileOutputStream(new File(OUTPUT_FILE))) {
      wb.write(out);
    }

  }

  private void createSheet(YearMonth ym) {
    XSSFSheet sheet = (XSSFSheet) wb.createSheet(ym.getMonthValue() + "月");
    sheet.setDefaultColumnWidth(3);
    IntStream.rangeClosed(0, 29).forEach(r -> {
      Row row = sheet.createRow(r);
      IntStream.rangeClosed(0, 31).forEach(col -> row.createCell(col));
    });
    sheet.addMergedRegion(new CellRangeAddress(0, 3, 0, 2));

    CreationHelper createHelper = wb.getCreationHelper();
    RichTextString title =
        createHelper.createRichTextString(ym.getYear() + "\r\n\r\n" + ym.getMonthValue() + "月");
    title.applyFont(0, 4, style.FONT_TITLE_YEAR);
    title.applyFont(8, title.length(), style.FONT_BOLD_18);
    Cell titleCell = sheet.getRow(0).getCell(0);
    titleCell.setCellStyle(style.CELL_STYLE_TITLE);
    titleCell.setCellValue(title);

    createHead(sheet, ym);
    createBody(sheet, ym);
  }

  private void createHead(XSSFSheet sheet, YearMonth ym) {

    YearMonth prev = ym.minusMonths(1);
    YearMonth next = ym.plusMonths(1);
    YearMonth afterNext = ym.plusMonths(2);

    printHeadDate(sheet, 3, prev);
    printHeadDate(sheet, 11, next);
    printHeadDate(sheet, 19, afterNext);
  }

  private void printHeadDate(Sheet sheet, int start, YearMonth ym) {

    Cell headMonth = sheet.getRow(0).getCell(start);
    CreationHelper createHelper = wb.getCreationHelper();
    RichTextString month = createHelper.createRichTextString(ym.getMonthValue() + "月");
    month.applyFont(style.FONT_BOLD);
    headMonth.setCellValue(month);

    Row dayOfWeekRow = sheet.getRow(1);

    IntStream.rangeClosed(0, 6).forEach(i -> {
      Cell dayOfWeekCell = dayOfWeekRow.createCell(start + i);
      dayOfWeekCell.setCellValue(style.WEEK_HEAD_STYLE_LIST.get(i));
      dayOfWeekCell.setCellStyle(style.CELL_STYLE_CENTER);
    });

    LocalDate date = ym.atDay(1);
    LocalDate endDate = ym.atEndOfMonth();

    Row row = sheet.getRow(2);
    int positionOfWeek = date.getDayOfWeek().getValue() - 1;
    while (!date.isAfter(endDate)) {

      RichTextString day = createHelper.createRichTextString(String.valueOf(date.getDayOfMonth()));

      if (isHoliday(date) || date.getDayOfWeek().equals(DayOfWeek.SUNDAY)) {
        day.applyFont(style.FONT_RED);
      } else if (date.getDayOfWeek().equals(DayOfWeek.SATURDAY)) {
        day.applyFont(style.FONT_BLUE);
      } else {
        day.applyFont(style.FONT_GREY);
      }
      Cell cell = row.getCell(start + positionOfWeek);
      cell.setCellValue(day);
      cell.setCellStyle(style.CELL_STYLE_CENTER);
      positionOfWeek++;
      if (date.getDayOfWeek() == DayOfWeek.SUNDAY) {
        int currentRow = row.getRowNum();
        row = sheet.getRow(currentRow + 1);
        positionOfWeek = 0;
      }
      date = date.plusDays(1);
    }
  }

  private void createBody(XSSFSheet sheet, YearMonth ym) {
    CreationHelper createHelper = wb.getCreationHelper();
    XSSFDrawing patriarch = sheet.createDrawingPatriarch();

    int mondayCol = 0;
    int tuesdayCol = 5;
    int wednesdayCol = 10;
    int thursdayCol = 15;
    int fridayCol = 20;
    int saturdayCol = 25;
    int sundayCol = 29;

    createDayOfWeekCell(sheet, createHelper, mondayCol, "月", style.FONT_BOLD_18);
    createDayOfWeekCell(sheet, createHelper, tuesdayCol, "火", style.FONT_BOLD_18);
    createDayOfWeekCell(sheet, createHelper, wednesdayCol, "水", style.FONT_BOLD_18);
    createDayOfWeekCell(sheet, createHelper, thursdayCol, "木", style.FONT_BOLD_18);
    createDayOfWeekCell(sheet, createHelper, fridayCol, "金", style.FONT_BOLD_18);
    createDayOfWeekCell(sheet, createHelper, saturdayCol, "土", style.FONT_BOLD_18_BLUE);
    createDayOfWeekCell(sheet, createHelper, sundayCol, "日", style.FONT_BOLD_18_RED);

    LocalDate date = ym.atDay(1);
    LocalDate endDate = ym.atEndOfMonth();

    int startRowIndex = 10;
    int rowCount = 0;
    while (!date.isAfter(endDate)) {
      int targetRow = startRowIndex + rowCount * 4;

      // 5週に収まらなかった場合
      if (rowCount == 5) {
        targetRow -= 3;
      }

      Row firstRow = sheet.getRow(targetRow);
      final Row row = firstRow;
      final int rc = rowCount;
      IntStream.rangeClosed(0, 31).forEach(idx -> {
        Cell c = row.getCell(idx);
        if (rc < 5) {
          c.setCellStyle(style.CELL_STYLE_TOP_BORDER);
        }
      });

      do {
        RichTextString day =
            createHelper.createRichTextString(String.valueOf(date.getDayOfMonth()));
        day.applyFont(style.FONT_BOLD_18);

        DayOfWeek w = date.getDayOfWeek();
        if (isHoliday(date)) {
          day.applyFont(style.FONT_BOLD_18_RED);
        }

        int dayCol = -1;
        switch (w) {
          case MONDAY:
            dayCol = mondayCol;
            break;
          case TUESDAY:
            dayCol = tuesdayCol;
            break;
          case WEDNESDAY:
            dayCol = wednesdayCol;
            break;
          case THURSDAY:
            dayCol = thursdayCol;
            break;
          case FRIDAY:
            dayCol = fridayCol;
            break;
          case SATURDAY:
            dayCol = saturdayCol;
            day.applyFont(style.FONT_BOLD_18_BLUE);
            if (isHoliday(date)) {
              day.applyFont(style.FONT_BOLD_18_RED);
            }
            break;
          case SUNDAY:
            dayCol = sundayCol;
            day.applyFont(style.FONT_BOLD_18_RED);
            break;
        }

        // 5週に収まらなかった場合前の週に押し込む
        if (rowCount == 5) {
          dayCol++;

          XSSFClientAnchor clientAnchor = patriarch.createAnchor(0, Units.EMU_PER_PIXEL * 10, 0,
              Units.EMU_PER_PIXEL * 10, (short) dayCol - 1, (short) targetRow - 1,
              (short) (dayCol + 1), (short) (targetRow + 1));
          XSSFSimpleShape shape = patriarch.createSimpleShape(clientAnchor);
          shape.setShapeType(ShapeTypes.LINE);
          shape.setLineWidth(1.0);
          // shape.setLineStyle(1);
          shape.setLineStyleColor(0, 0, 0);
          shape.getCTShape().getSpPr().getXfrm().setFlipV(true);
        }
        sheet.addMergedRegion(new CellRangeAddress(targetRow, targetRow + 1, dayCol, dayCol));
        row.getCell(dayCol).setCellValue(day);

        date = date.plusDays(1);
      } while (date.getDayOfWeek() != DayOfWeek.MONDAY && !date.isAfter(endDate));

      rowCount++;
    }
  }

  private void createDayOfWeekCell(XSSFSheet sheet, CreationHelper createHelper, int cellIndex,
      String dayOfWeek, Font font) {

    sheet.addMergedRegion(new CellRangeAddress(8, 9, cellIndex, cellIndex));
    RichTextString monday = createHelper.createRichTextString(dayOfWeek);
    monday.applyFont(font);
    Row week = sheet.getRow(8);
    week.getCell(cellIndex).setCellValue(monday);
    week.getCell(cellIndex).setCellStyle(style.CELL_STYLE_CENTER);
  }

  private boolean isHoliday(LocalDate date) {
    return holidays.is祝休日(date);
  }

  private void printSetting(int month) {
    // sheet index starts from 0
    int sheetIndex = month - 1;
    wb.setPrintArea(sheetIndex, 0, 31, 0, 29);
    Sheet sheet = wb.getSheetAt(sheetIndex);
    PrintSetup printSetup = sheet.getPrintSetup();
    printSetup.setLandscape(true);
    printSetup.setPaperSize(PrintSetup.TEN_BY_FOURTEEN_PAPERSIZE);
    printSetup.setFitWidth((short) 1);

    sheet.setMargin(Sheet.TopMargin, 0.1);
    sheet.setMargin(Sheet.BottomMargin, 0.1);
    sheet.setMargin(Sheet.LeftMargin, 0.3);
    sheet.setMargin(Sheet.RightMargin, 0.1);
    sheet.setMargin(Sheet.HeaderMargin, 0);
    sheet.setMargin(Sheet.FooterMargin, 0);
  }
}
