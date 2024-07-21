package com.onlyu.auto;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.time.LocalDate;
import java.time.Month;
import java.time.Year;
import java.time.temporal.ChronoUnit;

public class WeekEntry
{
    private int _index;
    private int _presentableIndex;
    private LocalDate _start;
    private LocalDate _end;
    private int _numberOfDays;
    private Sheet _sheet;
    private Row _contentRow;
    private Row _concludeRow;
    private Cell _period;
    private Cell _from;
    private Cell _to;
    private Cell _beginAt;
    private Cell _endAt;
    private Cell _totalTime;
    private Cell _taskDescription;
    private Cell _note;
    private Cell _totalTimeCalculated;
    private Cell _sent;

    WeekEntry(LocalDate start, LocalDate inclusive, Sheet sheet, int index)
    {
        if (start.isAfter(inclusive))
            return;
        switch (start.getDayOfWeek())
        {
            case MONDAY:
                _start = start;
                _end = start.plusDays(4);
                break;
            case TUESDAY:
                _start = start;
                _end = start.plusDays(3);
                break;
            case WEDNESDAY:
                _start = start;
                _end = start.plusDays(2);
                break;
            case THURSDAY:
                _start = start;
                _end = start.plusDays(1);
                break;
            case FRIDAY:
                _start = start;
                _end = start.plusDays(0);
                break;
            case SATURDAY:
                _start = start.plusDays(2);
                _end = start.plusDays(6);
                break;
            case SUNDAY:
                _start = start.plusDays(1);
                _end = start.plusDays(5);
                break;
        }
        if (_end.isAfter(inclusive))
            _end = inclusive;
        _numberOfDays = (int)_start.until(_end.plusDays(1), ChronoUnit.DAYS);
        _sheet = sheet;
        _index = index;
        _presentableIndex = _index + 1;
        _contentRow = _sheet.getRow(2 * _index + 6);
        _concludeRow = _sheet.getRow(2 * _index + 7);
        _period = _contentRow.getCell(0);
        _from = _contentRow.getCell(1);
        _to = _contentRow.getCell(2);
        _beginAt = _contentRow.getCell(3);
        _endAt = _contentRow.getCell(4);
        _totalTime = _contentRow.getCell(5);
        _taskDescription = _contentRow.getCell(6);
        _note = _contentRow.getCell(7);
        _totalTimeCalculated = _concludeRow.getCell(5);
        _sent = _concludeRow.getCell(6);
        if (_numberOfDays == 0)
        {
            _concludeRow.getCell(0).setCellValue("END OF MONTH, NO CONTENT HERE");
            _totalTimeCalculated.setCellValue("N/A");
            _sent.setCellValue("N/A");
            _period.setCellValue("N/A");
            _from.setCellValue("N/A");
            _to.setCellValue("N/A");
            _beginAt.setCellValue("N/A");
            _endAt.setCellValue("N/A");
            _totalTime.setCellValue("N/A");
            _taskDescription.setCellValue("N/A");
            _note.setCellValue("N/A");
        }
    }

    public static WeekEntry of(LocalDate start, LocalDate inclusive, Sheet sheet, int index)
    {
        return new WeekEntry(start, inclusive, sheet, index);
    }

    public Month getMonth()
    {
        return _start.getMonth();
    }

    public Year getYear()
    {
        return Year.of(_start.getYear());
    }

    public WeekEntry period(String period)
    {
        _period.setCellValue(period);
        return this;
    }

    public WeekEntry from(String from)
    {
        _from.setCellValue(from);
        return this;
    }

    public WeekEntry to(String to)
    {
        _to.setCellValue(to);
        return this;
    }

    public WeekEntry beginAt(String beginAt)
    {
        _beginAt.setCellValue(beginAt);
        return this;
    }

    public WeekEntry endAt(String endAt)
    {
        _endAt.setCellValue(endAt);
        return this;
    }

    public WeekEntry totalTime(String totalTime)
    {
        _totalTime.setCellValue(totalTime);
        return this;
    }

    public WeekEntry taskDescription(String taskDescription)
    {
        _taskDescription.setCellValue(taskDescription);
        return this;
    }

    public WeekEntry note(String note)
    {
        _note.setCellValue(note);
        return this;
    }

    public WeekEntry totalTimeCalculated(String totalTimeCalculated)
    {
        _totalTimeCalculated.setCellValue(totalTimeCalculated);
        return this;
    }

    public WeekEntry markAsSent()
    {
        _sent.setCellValue("SENT");
        return this;
    }

    public int getIndex()
    {
        return _index;
    }

    public int getPresentableIndex()
    {
        return _presentableIndex;
    }

    public Row getContentRow()
    {
        return _contentRow;
    }

    public Sheet getSheet()
    {
        return _sheet;
    }

    public Row getConcludeRow()
    {
        return _concludeRow;
    }

    public LocalDate getStartOfWeek()
    {
        return _start;
    }

    public int getNumberOfDays()
    {
        return _numberOfDays;
    }

    public LocalDate getEndOfWeek()
    {
        return _end;
    }

    public boolean hasBeenSent()
    {
        return !_sent.getStringCellValue().isBlank();
    }

    public boolean hasContent()
    {
        return !_totalTimeCalculated.getStringCellValue().isBlank();
    }

    public boolean isCurrent()
    {
        LocalDate now = LocalDate.now();
        return !now.isBefore(_start) && !now.isAfter(_end);
    }

    public boolean isPast()
    {
        LocalDate now = LocalDate.now();
        return now.isAfter(_end);
    }

    public boolean isFuture()
    {
        LocalDate now = LocalDate.now();
        return now.isBefore(_start);
    }
}
