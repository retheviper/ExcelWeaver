package com.retheviper.excel.definition;

import lombok.AccessLevel;
import lombok.Getter;
import lombok.RequiredArgsConstructor;
import lombok.ToString;

/**
 * Definition of sheet.
 */
@Getter
@ToString
@RequiredArgsConstructor(access = AccessLevel.PACKAGE)
public class SheetDef {

    /**
     * Definition of row, which means header.
     */
    private final RowDef rowDef;

    /**
     * Sheet's name.
     */
    private final String name;

    /**
     * Sheet's index.
     */
    private final int index;

    /**
     * Row index which data part begins.
     */
    private final int dataStartIndex;

    /**
     * Object to map with this Sheet. The object will be considered as sheet, and fields are considered as cells.
     */
    private final Class<?> dataClass;
}
