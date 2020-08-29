# ExcelWeaver

A simple library for manipulate excel file in java.

## TL:DR

`*.xlsx` → `java.util.List`

`java.util.List` → `*.xlsx`

## To build

`./gradlew jar`

## To Use

### Make definition of sheet

```
@Sheet(dataStartIndex = 2)
public class Contract {

    @Column(position = "B")
    private String name;

    @Column(position = "C")
    private String cellPhone;

    @Column(position = "D")
    private int postCode;
}
```

### Add sheet to book

```
BookDef bookDef = new BookDef();
bookDef.addSheet(Contract.class);
```

### Read file

```
List<Contract> list;
try (BookWorker worker = bookDef.openBook(OUTPUT_FILE_PATH)) {
    list = worker.read(Contract.class);
} catch (IOException e) {
    // ...
}
```

### Write file

```
try (BookWorker worker = bookDef.openBook(TEMPLATE_FILE_PATH, OUTPUT_FILE_PATH)) {
    worker.write(data);
} catch (IOException e) {
    // ...
}
```