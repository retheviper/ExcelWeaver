# ExcelWeaver

![exceweaver](./misc/excelweaver.jpg)

A simple library for manipulate excel file in java.

## TL:DR

`*.xlsx` → `java.util.List`

`java.util.List` → `*.xlsx`

## To build

`./gradlew shadowJar`

(And Jar will be found in `build/libs`)

## To Use

### Make definition of sheet

```java
@Sheet(dataStartIndex = 2)
public class Contract { // Class name will be sheets name

    @Column(position = "B")
    private String name;

    @Column(position = "C")
    private String cellPhone;

    @Column(position = "D")
    private int postCode;
}
```

### Make definition of book

```java
// Create from array or list of SheetDef classes
BookDef bookDef = BookDef.of(Contract.class, Message.class);

// add more sheet
bookDef.addSheet(Salary.class);
```

### Read file

```java
List<Contract> list;
try (BookWorker worker = bookDef.openBook(OUTPUT_FILE_PATH)) {
    list = worker.read(Contract.class);
} catch (IOException e) {
    // ...
}
```

### Write file

```java
List<Contract> data = ...
try (BookWorker worker = bookDef.openBook(TEMPLATE_FILE_PATH, OUTPUT_FILE_PATH)) {
    worker.write(data);
} catch (IOException e) {
    // ...
}
```