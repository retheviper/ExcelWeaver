package com.retheviper.excel.worker;

import com.retheviper.excel.definition.BookDef;
import com.retheviper.excel.definition.SheetDef;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import java.util.function.Predicate;

/**
 * Worker for manipulate excel file.
 */
public class BookWorker implements AutoCloseable {

    /**
     * Definition of book.
     */
    private final BookDef bookDef;

    /**
     * Apache POI's workbook class.
     */
    private final Workbook workbook;

    /**
     * Path for output.
     */
    private Path outputPath = null;

    /**
     * Create BookWorker for read data.
     *
     * @param bookDef
     * @param inputPath
     * @throws IOException
     */
    public BookWorker(final BookDef bookDef, final String inputPath) throws IOException {
        this.bookDef = bookDef;
        this.workbook = WorkbookFactory.create(new File(inputPath));
    }

    /**
     * Create BookWorker for write Data.
     *
     * @param bookDef
     * @param templatePath
     * @param outputPath
     * @throws IOException
     */
    public BookWorker(final BookDef bookDef, final String templatePath, final String outputPath) throws IOException {
        this.bookDef = bookDef;
        this.workbook = WorkbookFactory.create(new File(templatePath));
        this.outputPath = Paths.get(outputPath);
    }

    /**
     * Get {@link SheetWorker}.
     *
     * @param sheetName
     * @return SheetWorker
     */
    public SheetWorker getSheetWorker(final String sheetName) {
        return searchSheetDef(sheetDef -> Objects.equals(sheetDef.getName(), sheetName)).orElseThrow(NullPointerException::new);
    }

    /**
     * Get {@link SheetWorker}.
     *
     * @param index
     * @return SheetWorker
     */
    public SheetWorker getSheetWorker(final int index) {
        return searchSheetDef(sheetDef -> index == sheetDef.getIndex()).orElseThrow(NullPointerException::new);
    }

    /**
     * Get {@link SheetWorker}.
     *
     * @param dataClass
     * @return SheetWorker
     */
    public SheetWorker getSheetWorker(final Class<?> dataClass) {
        return searchSheetDef(sheetDef -> Objects.equals(sheetDef.getDataClass(), dataClass)).orElseThrow(NullPointerException::new);
    }

    /**
     * Search {@link SheetDef} with condition and map to {@link SheetWorker}.
     *
     * @param condition
     * @return
     */
    private Optional<SheetWorker> searchSheetDef(final Predicate<SheetDef> condition) {
        return this.bookDef.getSheetDefs().stream().filter(condition).findAny().map(this::getSheetWorker);
    }

    /**
     * Get {@link SheetWorker}.
     *
     * @param sheetDef
     * @return
     */
    private SheetWorker getSheetWorker(final SheetDef sheetDef) {
        return new SheetWorker(sheetDef, sheetDef.getIndex() != -1 ? this.workbook.getSheetAt(sheetDef.getIndex()) : this.workbook.getSheet(sheetDef.getName()));
    }

    /**
     * Write data into sheet.
     *
     * @param data
     * @param <T>
     */
    public <T> void write(final List<T> data) {
        getSheetWorker(data.get(0).getClass()).listToSheet(data);
    }

    /**
     * Read data from sheet as list.
     *
     * @param sheetName
     * @param <T>
     * @return
     */
    public <T> List<T> read(final String sheetName) {
        return getSheetWorker(sheetName).sheetToList();
    }

    /**
     * Read data from sheet as list.
     *
     * @param index
     * @param <T>
     * @return
     */
    public <T> List<T> read(final int index) {
        return getSheetWorker(index).sheetToList();
    }

    /**
     * Read data from sheet as list.
     *
     * @param dataClass
     * @param <T>
     * @return
     */
    public <T> List<T> read(final Class<T> dataClass) {
        return getSheetWorker(dataClass).sheetToList();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void close() throws IOException {
        if (this.outputPath != null) {
            if (Files.notExists(outputPath.getParent())) {
                Files.createDirectories(outputPath.getParent());
            }
            try (final OutputStream stream = new BufferedOutputStream(Files.newOutputStream(outputPath))) {
                workbook.write(stream);
            }
        } else {
            this.workbook.close();
        }
    }
}
