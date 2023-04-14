# Read_Excel_Poi
PoiでExcelファイルを読み込む。

- [Read\_Excel\_Poi](#read_excel_poi)
  - [実行](#実行)
  - [src](#src)
    - [build.gradle](#buildgradle)
    - [ittimfn/sample/poi/read/App.java](#ittimfnsamplepoireadappjava)
    - [ittimfn/sample/poi/read/controller/ReadExcelController.java](#ittimfnsamplepoireadcontrollerreadexcelcontrollerjava)
    - [ittimfn/sample/poi/read/model/ExcelModel.java](#ittimfnsamplepoireadmodelexcelmodeljava)

## 実行

``` bash
./gradlew run --args="$(pwd)/SampleExcel.xlsx"
```

## src

### build.gradle

``` groovy
/*
 * This file was generated by the Gradle 'init' task.
 *
 * This generated file contains a sample Java application project to get you started.
 * For more details take a look at the 'Building Java & JVM projects' chapter in the Gradle
 * User Manual available at https://docs.gradle.org/7.6/userguide/building_java_projects.html
 */

plugins {
    id 'java'
    // Apply the application plugin to add support for building a CLI application in Java.
    id 'application'
    id 'eclipse'
}

repositories {
    // Use Maven Central for resolving dependencies.
    mavenCentral()
}

dependencies {

    implementation 'org.apache.logging.log4j:log4j-api:2.20.0'
    implementation 'org.apache.logging.log4j:log4j-core:2.20.0'

    implementation 'org.apache.poi:poi:5.2.3'
    implementation 'org.apache.poi:poi-ooxml:5.2.3'

	compileOnly 'org.projectlombok:lombok:1.18.26'
	annotationProcessor 'org.projectlombok:lombok:1.18.26'
    
    // Use JUnit Jupiter for testing.
    testImplementation 'org.junit.jupiter:junit-jupiter:5.9.1'

    // This dependency is used by the application.
    implementation 'com.google.guava:guava:31.1-jre'
}

application {
    // Define the main class for the application.
    mainClass = 'ittimfn.sample.poi.read.App'
}

tasks.named('test') {
    // Use JUnit Platform for unit tests.
    useJUnitPlatform()
}

```

### ittimfn/sample/poi/read/App.java

``` java
/*
 * This Java source file was generated by the Gradle 'init' task.
 */
package ittimfn.sample.poi.read;

import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;

import ittimfn.sample.poi.read.controller.ReadExcelController;

public class App {

    private ReadExcelController controller;

    public void exec(String[] args) throws EncryptedDocumentException, FileNotFoundException, IOException {
        int argsIndex = 0;
        String excelFilePath = args[argsIndex++];
        
        this.controller = new ReadExcelController(excelFilePath);
        System.out.println(this.controller.getModel());
        this.controller.close();
    }

    public static void main(String[] args) throws EncryptedDocumentException, FileNotFoundException, IOException {
        new App().exec(args);
    }
}

```

### ittimfn/sample/poi/read/controller/ReadExcelController.java

``` java
package ittimfn.sample.poi.read.controller;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import ittimfn.sample.poi.read.model.ExcelModel;
import lombok.Data;

@Data
public class ReadExcelController {

    private FileInputStream stream;

    private Workbook workbook;
    private Sheet sheet;

    public ReadExcelController(String filepath) throws EncryptedDocumentException, FileNotFoundException, IOException {
        this.stream = new FileInputStream(filepath);
        this.workbook = WorkbookFactory.create(this.stream);
        this.sheet = workbook.getSheet("Sheet1");
    }

    public ExcelModel getModel() {
        ExcelModel model = new ExcelModel();

        Row row = sheet.getRow(0);
        Cell cell = row.getCell(0);

        model.setCellValue(cell.getStringCellValue());

        return model;
    }

    public void close() throws IOException {
        this.workbook.close();
        this.stream.close();
    }


}

```

### ittimfn/sample/poi/read/model/ExcelModel.java

``` java
package ittimfn.sample.poi.read.model;

import lombok.Data;

@Data
public class ExcelModel {
    private String cellValue;
}

```

