package com.retheviper.excel.definition;

import com.retheviper.excel.definition.annotation.Column;
import com.retheviper.excel.definition.annotation.Sheet;
import com.retheviper.excel.definition.type.DataType;
import com.retheviper.excel.worker.BookWorker;
import lombok.Getter;
import lombok.ToString;
import org.apache.poi.ss.util.CellAddress;

import java.io.IOException;
import java.lang.reflect.Field;
import java.time.LocalDateTime;
import java.util.*;
import java.util.stream.Collectors;

/**
 * Definition of entire excel file, book.
 */
@Getter
@ToString
public class BookDef {

    private final List<SheetDef> sheetDefs = new ArrayList<>();

    /**
     * Add definition of sheet based on current data classes. The data class must be contains {@link Sheet} and {@link Column}.
     *
     * @param dataClasses
     * @return BookDef
     * @throws Exception
     */
    public static BookDef of(final List<Class<?>> dataClasses) throws Exception {
        final BookDef bookDef = new BookDef();
        for (Class<?> dataClass : dataClasses) {
            bookDef.addSheet(dataClass);
        }
        return bookDef;
    }

    /**
     * Add definition of sheet based on current data classes. The data class must be contains {@link Sheet} and {@link Column}.
     *
     * @param dataClasses
     * @return BookDef
     * @throws Exception
     */
    public static BookDef of(final Class<?>... dataClasses) throws Exception {
        final BookDef bookDef = new BookDef();
        for (Class<?> dataClass : dataClasses) {
            bookDef.addSheet(dataClass);
        }
        return bookDef;
    }

    /**
     * Add definition of sheet based on current data class. The data class must be contains {@link Sheet} and {@link Column}.
     *
     * @param dataClass
     * @throws Exception
     */
    public void addSheet(final Class<?> dataClass) throws Exception {
        this.sheetDefs.add(dataClassToSheetDef(dataClass));
    }

    /**
     * Remove definition of sheet based on current data class.
     *
     * @param dataClass
     */
    public void removeSheet(final Class<?> dataClass) {
        this.sheetDefs.removeIf(sheetDef -> Objects.equals(sheetDef.getDataClass(), dataClass));
    }

    /**
     * Find and returns definition of sheet based on current data class.
     * If null, {@link NullPointerException} will be thrown.
     *
     * @param dataClass
     * @return
     */
    public SheetDef getSheetDef(final Class<?> dataClass) {
        return this.sheetDefs.stream().filter(sheetDef -> Objects.equals(sheetDef.getDataClass(), dataClass)).findAny().orElseThrow(NullPointerException::new);
    }

    /**
     * Create BookWorker for read data from excel file.
     *
     * @param path
     * @return
     * @throws IOException
     */
    public BookWorker openBook(final String path) throws IOException {
        return new BookWorker(this, path);
    }

    /**
     * Create BookWorker for write data with template excel file.
     *
     * @param templatePath
     * @param outputPath
     * @return
     * @throws IOException
     */
    public BookWorker openBook(final String templatePath, final String outputPath) throws IOException {
        return new BookWorker(this, templatePath, outputPath);
    }

    /**
     * Create {@link SheetDef} from data class.
     *
     * @param dataClass
     * @return
     * @throws Exception
     */
    private SheetDef dataClassToSheetDef(final Class<?> dataClass) throws Exception {
        if (!dataClass.isAnnotationPresent(Sheet.class)) {
            throw new IllegalArgumentException(String.format("Missing annotation: %s", dataClass.getSimpleName()));
        }
        final RowDef rowDef = new RowDef(Arrays.stream(dataClass.getDeclaredFields()).filter(field -> field.isAnnotationPresent(Column.class)).map(this::fieldToCellDef).collect(Collectors.toList()));
        final Sheet sheet = dataClass.getAnnotation(Sheet.class);
        return new SheetDef(rowDef, !sheet.name().isBlank() ? sheet.name() : dataClass.getSimpleName(), sheet.dataStartIndex(), dataClass);
    }

    /**
     * Create {@link CellDef} from field.
     *
     * @param field
     * @return
     */
    private CellDef fieldToCellDef(final Field field) {
        field.setAccessible(true);
        final Column column = field.getAnnotation(Column.class);
        return new CellDef(!column.name().isBlank() ? column.name() : field.getName(), defineFieldType(field.getType()), new CellAddress(column.position() + 1).getColumn());
    }


    /**
     * Define field's data type.
     *
     * @param fieldType
     * @return
     */
    private DataType defineFieldType(final Class<?> fieldType) {
        if (Integer.TYPE.equals(fieldType) || Integer.class.equals(fieldType)) {
            return DataType.INTEGER;
        } else if (Long.TYPE.equals(fieldType) || Long.class.equals(fieldType)) {
            return DataType.LONG;
        } else if (Double.TYPE.equals(fieldType) || Double.class.equals(fieldType)) {
            return DataType.DOUBLE;
        } else if (Date.class.equals(fieldType)) {
            return DataType.DATE;
        } else if (LocalDateTime.class.equals(fieldType)) {
            return DataType.LOCAL_DATE_TIME;
        } else if (Boolean.TYPE.equals(fieldType) || Boolean.class.equals(fieldType)) {
            return DataType.BOOLEAN;
        } else {
            return DataType.STRING;
        }
    }
}
