package com.retheviper.excel.worker;

import com.retheviper.excel.definition.CellDef;
import com.retheviper.excel.definition.RowDef;
import org.apache.poi.ss.usermodel.Row;

import java.util.List;
import java.util.stream.Collectors;

/**
 * Worker for manipulate row.
 */
class RowWorker {

    /**
     * Definition of row.
     */
    private final RowDef rowDef;

    /**
     * List of {@link CellWorker}.
     */
    private final List<CellWorker> cellWorkers;

    /**
     * Apache POI's row class.
     */
    private final Row row;

    /**
     * Current index of cell. Increases when {@link RowWorker#nextCell()} called.
     */
    private int currentCellIndex;

    /**
     * Constructor.
     *
     * @param rowDef
     * @param row
     */
    protected RowWorker(final RowDef rowDef, final Row row) {
        this.rowDef = rowDef;
        this.row = row;
        this.cellWorkers = rowDef.getCellDefs().stream().map(cell -> getCellWorker(cell.getCellIndex())).collect(Collectors.toList());
        this.currentCellIndex = rowDef.getCellDefs().get(0).getCellIndex();
    }

    /**
     * Get next {@link CellWorker}.
     *
     * @return
     */
    public CellWorker nextCell() {
        return getCellWorker(this.currentCellIndex++);
    }

    /**
     * Get {@link CellWorker} with current index.
     *
     * @param cellIndex
     * @return
     */
    public CellWorker getCellWorker(final int cellIndex) {
        return new CellWorker(getCellDef(cellIndex), this.row.getCell(cellIndex));
    }

    /**
     * Get {@link CellDef} with current index.
     *
     * @param cellIndex
     * @return
     */
    private CellDef getCellDef(final int cellIndex) {
        return this.rowDef.getCellDefs().stream().filter(cell -> cell.getCellIndex() == cellIndex).findFirst()
                .orElseThrow(NullPointerException::new);
    }

    /**
     * Read row's value and write into object's field.
     *
     * @param object
     * @param <T>
     */
    public <T> void rowToObj(final T object) {
        for (CellWorker cellWorker : cellWorkers) {
            cellWorker.cellToObj(object);
        }
    }

    /**
     * Read object's field value and write into cell's field.
     *
     * @param object
     * @param <T>
     */
    public <T> void objToRow(final T object) {
        for (CellWorker cellWorker : cellWorkers) {
            cellWorker.objToCell(object);
        }
    }
}
