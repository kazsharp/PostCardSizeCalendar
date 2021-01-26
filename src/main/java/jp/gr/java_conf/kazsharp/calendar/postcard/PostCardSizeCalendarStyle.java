package jp.gr.java_conf.kazsharp.calendar.postcard;

import java.util.List;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

public class PostCardSizeCalendarStyle {

  public Font FONT_TITLE_YEAR;
  public Font FONT_BOLD_18;
  public Font FONT_BOLD_18_BLUE;
  public Font FONT_BOLD_18_RED;
  public Font FONT_BOLD;
  public Font FONT_RED;
  public Font FONT_BLUE;
  public Font FONT_GREY;

  public CellStyle CELL_STYLE_TITLE;
  public CellStyle CELL_STYLE_CENTER;
  public CellStyle CELL_STYLE_TOP_BORDER;
  public CellStyle CELL_STYLE_ALIGNMENT_TOP;

  public RichTextString HEAD_MONDAY;
  public RichTextString HEAD_TUESDAY;
  public RichTextString HEAD_WEDNESDAY;
  public RichTextString HEAD_THURSDAY;
  public RichTextString HEAD_FRIDAY;
  public RichTextString HEAD_SATURDAY;
  public RichTextString HEAD_SUNDAY;

  public List<RichTextString> WEEK_HEAD_STYLE_LIST;

  public PostCardSizeCalendarStyle(Workbook wb) {
    FONT_TITLE_YEAR = wb.createFont();
    FONT_TITLE_YEAR.setFontName("メイリオ");
    FONT_TITLE_YEAR.setBold(true);
    FONT_TITLE_YEAR.setColor(IndexedColors.BLUE.getIndex());
    FONT_TITLE_YEAR.setFontHeightInPoints((short) 14);

    FONT_BOLD_18 = wb.createFont();
    FONT_BOLD_18.setFontName("メイリオ");
    FONT_BOLD_18.setBold(true);
    FONT_BOLD_18.setFontHeightInPoints((short) 18);

    FONT_BOLD_18_BLUE = wb.createFont();
    FONT_BOLD_18_BLUE.setFontName("メイリオ");
    FONT_BOLD_18_BLUE.setBold(true);
    FONT_BOLD_18_BLUE.setFontHeightInPoints((short) 18);
    FONT_BOLD_18_BLUE.setColor(IndexedColors.BLUE.getIndex());

    FONT_BOLD_18_RED = wb.createFont();
    FONT_BOLD_18_RED.setFontName("メイリオ");
    FONT_BOLD_18_RED.setBold(true);
    FONT_BOLD_18_RED.setFontHeightInPoints((short) 18);
    FONT_BOLD_18_RED.setColor(IndexedColors.RED.getIndex());

    FONT_BOLD = wb.createFont();
    FONT_BOLD.setFontName("メイリオ");
    FONT_BOLD.setBold(true);

    FONT_RED = wb.createFont();
    FONT_RED.setFontName("メイリオ");
    FONT_RED.setColor(IndexedColors.RED.getIndex());

    FONT_BLUE = wb.createFont();
    FONT_BLUE.setFontName("メイリオ");
    FONT_BLUE.setColor(IndexedColors.BLUE.getIndex());

    FONT_GREY = wb.createFont();
    FONT_GREY.setFontName("メイリオ");
    FONT_GREY.setColor(IndexedColors.GREY_50_PERCENT.getIndex());

    CELL_STYLE_TITLE = wb.createCellStyle();
    CELL_STYLE_TITLE.setWrapText(true);
    CELL_STYLE_TITLE.setAlignment(HorizontalAlignment.LEFT);
    CELL_STYLE_TITLE.setVerticalAlignment(VerticalAlignment.CENTER);

    CELL_STYLE_CENTER = wb.createCellStyle();
    CELL_STYLE_CENTER.setAlignment(HorizontalAlignment.CENTER);

    CELL_STYLE_TOP_BORDER = wb.createCellStyle();
    CELL_STYLE_TOP_BORDER.setBorderTop(BorderStyle.THIN);
    CELL_STYLE_TOP_BORDER.setAlignment(HorizontalAlignment.CENTER);
    CELL_STYLE_TOP_BORDER.setVerticalAlignment(VerticalAlignment.TOP);

    CELL_STYLE_ALIGNMENT_TOP = wb.createCellStyle();
    CELL_STYLE_ALIGNMENT_TOP.setVerticalAlignment(VerticalAlignment.TOP);

    HEAD_MONDAY = wb.getCreationHelper().createRichTextString("月");
    HEAD_MONDAY.applyFont(FONT_GREY);
    HEAD_TUESDAY = wb.getCreationHelper().createRichTextString("火");
    HEAD_TUESDAY.applyFont(FONT_GREY);
    HEAD_WEDNESDAY = wb.getCreationHelper().createRichTextString("水");
    HEAD_WEDNESDAY.applyFont(FONT_GREY);
    HEAD_THURSDAY = wb.getCreationHelper().createRichTextString("木");
    HEAD_THURSDAY.applyFont(FONT_GREY);
    HEAD_FRIDAY = wb.getCreationHelper().createRichTextString("金");
    HEAD_FRIDAY.applyFont(FONT_GREY);
    HEAD_SATURDAY = wb.getCreationHelper().createRichTextString("土");
    HEAD_SATURDAY.applyFont(FONT_BLUE);
    HEAD_SUNDAY = wb.getCreationHelper().createRichTextString("日");
    HEAD_SUNDAY.applyFont(FONT_RED);

    WEEK_HEAD_STYLE_LIST = List.of(HEAD_MONDAY, HEAD_TUESDAY, HEAD_WEDNESDAY, HEAD_THURSDAY, HEAD_FRIDAY,
        HEAD_SATURDAY, HEAD_SUNDAY);

  }
}
