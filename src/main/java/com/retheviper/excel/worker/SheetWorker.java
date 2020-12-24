package com.retheviper.excel.worker;

import com.retheviper.excel.definition.RowDef;
import com.retheviper.excel.definition.SheetDef;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

/**
 * Worker for manipulate sheet.
 */
class SheetWorker {

    /**
     * Definition of sheet.
     */
    private final SheetDef sheetDef;

    /**
     * Apache POI's sheet class.
     */
    private final Sheet sheet;

    /**
     * Current index of row. Increases when {@link SheetWorker#nextRow()} called.
     */
    private int currentRowIndex;

    /**
     * Constructor.
     *
     * @param sheetDef
     * @param sheet
     */
    protected SheetWorker(final SheetDef sheetDef, final Sheet sheet) {
        this.sheetDef = sheetDef;
        this.sheet = sheet;
        this.currentRowIndex = sheetDef.getDataStartIndex();
    }

    /**
     * Get next {@link RowWorker}.
     *
     * @return
     */
    public RowWorker nextRow() {
        return getRowWorker(this.currentRowIndex++);
    }

    /**
     * Get {@link RowWorker} with current index.
     *
     * @param rowIndex
     * @return
     */
    public RowWorker getRowWorker(final int rowIndex) {
        return new RowWorker(getRowDef(), getRow(rowIndex));
    }

    /**
     * Get Row.
     * If null, create new row and copies style of cells from first row.
     *
     * @param rowIndex
     * @return
     */
    private Row getRow(final int rowIndex) {
        final Row row = this.sheet.getRow(rowIndex) != null ? this.sheet.getRow(rowIndex) : this.sheet.createRow(rowIndex);
        final Row firstRow = this.sheet.getRow(this.sheetDef.getDataStartIndex());
        row.setRowStyle(firstRow.getRowStyle());
        if (row.getFirstCellNum() < this.getRowDef().getCellDefs().get(0).getCellIndex()) {
            firstRow.forEach(cell -> row.createCell(cell.getColumnIndex(), cell.getCellType()).setCellStyle(cell.getCellStyle()));
        }
        return row;
    }

    /**
     * Get definition of row.
     *
     * @return
     */
    private RowDef getRowDef() {
        return this.sheetDef.getRowDef();
    }

    /**
     * Read objects value and write into sheet.
     *
     * @param data
     * @param <T>
     */
    public <T> void listToSheet(List<T> data) {
        data.forEach(object -> nextRow().objToRow(object));
    }

    /**
     * Read sheet's value and write into objects.
     *
     * @param <T>
     * @return
     */
    @SuppressWarnings("unchecked")
    public <T> List<T> sheetToList() {
        return IntStream.rangeClosed(sheetDef.getDataStartIndex(), this.sheet.getLastRowNum()).mapToObj(index -> {
            try {
                final Object object = this.sheetDef.getDataClass().getDeclaredConstructor().newInstance();
                getRowWorker(index).rowToObj(object);
                return (T) object;
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        }).collect(Collectors.toUnmodifiableList());
    }
}
