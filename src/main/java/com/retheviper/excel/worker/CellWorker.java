package com.retheviper.excel.worker;

import com.retheviper.excel.definition.CellDef;
import lombok.AccessLevel;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;

import java.lang.reflect.Field;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Date;

/**
 * Worker for manipulate cell.
 */
@RequiredArgsConstructor(access = AccessLevel.PACKAGE)
class CellWorker {

    /**
     * Definition of cell.
     */
    private final CellDef cellDef;

    /**
     * Apache POI's cell class.
     */
    private final Cell cell;

    /**
     * Read cell's value and write into object's field.
     *
     * @param object
     * @param <T>
     * @return
     */
    public <T> T cellToObj(final T object) {
        final Field field = getObjectField(object);
        try {
            switch (cellDef.getType()) {
                case INTEGER:
                    field.set(object, (int) this.cell.getNumericCellValue());
                    break;
                case LONG:
                    field.set(object, (long) this.cell.getNumericCellValue());
                    break;
                case DOUBLE:
                    field.set(object, this.cell.getNumericCellValue());
                    break;
                case STRING:
                    field.set(object, this.cell.getStringCellValue().isBlank() ? null : this.cell.getStringCellValue());
                    break;
                case BOOLEAN:
                    field.set(object, this.cell.getBooleanCellValue());
                    break;
                case LOCAL_DATE_TIME:
                    field.set(object, this.cell.getLocalDateTimeCellValue());
                    break;
                case LOCAL_DATE:
                    field.set(object, this.cell.getLocalDateTimeCellValue().toLocalDate());
                    break;
                case DATE:
                    field.set(object, DateUtil.getJavaDate(this.cell.getNumericCellValue()));
                    break;
            }
        } catch (Exception e) {
            return null;
        }
        return object;
    }

    /**
     * Read object's field value and write into cell's field.
     *
     * @param object
     * @param <T>
     * @return
     */
    public <T> void objToCell(final T object) {
        final Field field = getObjectField(object);
        try {
            final Object objectValue = field.get(object);
            switch (cellDef.getType()) {
                case INTEGER:
                    this.cell.setCellValue((int) objectValue);
                    break;
                case LONG:
                    this.cell.setCellValue((long) objectValue);
                    break;
                case DOUBLE:
                    this.cell.setCellValue((double) objectValue);
                    break;
                case STRING:
                    this.cell.setCellValue((String) objectValue);
                    break;
                case BOOLEAN:
                    this.cell.setCellValue((boolean) objectValue);
                    break;
                case LOCAL_DATE_TIME:
                    this.cell.setCellValue((LocalDateTime) objectValue);
                    break;
                case LOCAL_DATE:
                    this.cell.setCellValue((LocalDate) objectValue);
                    break;
                case DATE:
                    this.cell.setCellValue(DateUtil.getExcelDate((Date) objectValue));
                    break;
            }
        } catch (Exception e) {
            return;
        }
    }

    /**
     * Get field from object.
     *
     * @param object
     * @param <T>
     * @return
     */
    private <T> Field getObjectField(final T object) {
        try {
            final Field field = object.getClass().getDeclaredField(this.cellDef.getName());
            field.setAccessible(true);
            return field;
        } catch (NoSuchFieldException e) {
            throw new RuntimeException(e);
        }
    }
}
